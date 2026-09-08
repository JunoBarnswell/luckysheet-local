package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.NativeKernelIntegrationTestSupport;
import com.xc.luckysheet.server.contract.CopyWorkbookRequest;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.ShareCreateRequest;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookManifestEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookSourceArtifactEntityRepository;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.transaction.PlatformTransactionManager;
import org.springframework.transaction.support.TransactionTemplate;
import org.springframework.test.context.TestPropertySource;

import java.time.Instant;
import java.util.List;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.*;

@TestPropertySource(properties = "DATABASE_URL=jdbc:h2:mem:native_catalog;DB_CLOSE_DELAY=-1")
class WorkbookCatalogServiceTest extends NativeKernelIntegrationTestSupport {
    @Autowired private WorkbookCatalogService catalog;
    @Autowired private WorkbookOperationService operations;
    @Autowired private WorkbookEntityRepository workbooks;
    @Autowired private WorkbookManifestEntityRepository manifests;
    @Autowired private WorkbookSourceArtifactEntityRepository artifacts;
    @Autowired private KernelHostClient kernel;
    @Autowired private GuestShareService shares;
    @Autowired private ObjectMapper mapper;
    @Autowired private PlatformTransactionManager transactionManager;

    @Test
    void createPublishesNativeManifestAndReopensAfterProcessRestart() {
        var created = catalog.create(new CreateWorkbookRequest("native-catalog", "Catalog", null, null, null, null, null), "owner");
        assertEquals(0, created.revision());
        assertEquals("WorkbookManifest", created.manifest().path("schema").asText());
        assertEquals(11, created.manifest().path("version").asInt());
        assertFalse(created.manifest().has("cells"));
        assertTrue(manifests.findByUnitIdAndRevision("native-catalog", 0).isPresent());
        kernel.close();
        var reopened = operations.open("native-catalog", "owner");
        assertEquals(created.manifest(), reopened.manifest());
        assertEquals(created.checksum(), reopened.checksum());
    }

    @Test
    void invalidSheetManifestIsRejectedWithoutWorkbookOrManifestPersistence() throws Exception {
        var sheets = mapper.readTree("""
                [{"sheetId":"sheet-1","name":"Invalid","rowCount":0,"columnCount":26,"metadata":{}}]
                """);
        var request = new CreateWorkbookRequest("invalid-native-catalog", "Invalid", sheets, null, null, null, null);
        var error = assertThrows(KernelHostException.class, () -> catalog.create(request, "owner"));
        assertEquals("SHEET_INVALID", error.code());
        assertFalse(workbooks.existsById(request.unitId()));
        assertTrue(manifests.findByUnitIdAndRevision(request.unitId(), 0).isEmpty());
    }

    @Test
    void createCommitsTemplateCellsThroughTheCanonicalOperationChain() throws Exception {
        var sheets = mapper.readTree("""
                [{"sheetId":"sheet-1","name":"Template","rowCount":1000,"columnCount":26,"metadata":{}}]
                """);
        var params = mapper.createObjectNode().put("sheetId", "sheet-1").put("row", 2).put("column", 3);
        params.putObject("value").put("value", "created through Rust");

        var created = catalog.create(new CreateWorkbookRequest(
                "native-template-create", "Template", sheets, null, null, null,
                List.of(new OperationMutation("cell.set", "sheet-1", params))), "template-owner");

        assertEquals(1, created.revision());
        assertEquals(1, created.manifest().path("revision").asLong());
        assertEquals(1, created.manifest().path("pages").size());
        assertTrue(manifests.findByUnitIdAndRevision("native-template-create", 0).isPresent());
        assertTrue(manifests.findByUnitIdAndRevision("native-template-create", 1).isPresent());
        kernel.call("open", mapper.createObjectNode().set("manifest", created.manifest()));
        var load = mapper.createObjectNode().put("unitId", created.unitId()).put("revision", created.revision());
        load.set("page", operations.page(created.unitId(), created.revision(), "sheet-1", 0, 0, "template-owner"));
        kernel.call("page.load", load);
        var address = mapper.createObjectNode().put("sheetId", "sheet-1").put("row", 2).put("column", 3);
        var read = kernel.call("cell.get", mapper.createObjectNode()
                .put("unitId", "native-template-create").put("revision", 1).set("address", address));
        assertEquals("created through Rust", read.path("cell").path("value").asText());
    }

    @Test
    void rejectedTemplateCellRollsBackWorkbookAndEveryCheckpoint() throws Exception {
        var sheets = mapper.readTree("""
                [{"sheetId":"sheet-1","name":"Template","rowCount":1000,"columnCount":26,"metadata":{}}]
                """);
        var invalid = mapper.createObjectNode().put("row", 0).put("column", 0);
        invalid.putObject("value").put("value", "missing sheet identity");
        var request = new CreateWorkbookRequest(
                "invalid-template-create", "Invalid template", sheets, null, null, null,
                List.of(new OperationMutation("cell.set", "sheet-1", invalid)));

        assertThrows(KernelHostException.class, () -> catalog.create(request, "template-owner"));
        assertFalse(workbooks.existsById(request.unitId()));
        assertTrue(manifests.findByUnitIdAndRevision(request.unitId(), 0).isEmpty());
        assertTrue(manifests.findByUnitIdAndRevision(request.unitId(), 1).isEmpty());
    }

    @Test
    void nativeCopyHasIndependentIdentityAndCellsAcrossPersistedReopen() throws Exception {
        var sheets = mapper.readTree("""
                [{"sheetId":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,"metadata":{}}]
                """);
        catalog.create(new CreateWorkbookRequest("native-copy-source", "Source", sheets, null, null, null, null), "copy-owner");
        catalog.exportArtifact("native-copy-source", 0, "source.xlsx", "xlsx", "copy-owner");
        var params = mapper.createObjectNode().put("sheetId", "sheet-1").put("row", 0).put("column", 0);
        params.putObject("value").put("value", 42);
        var write = new OperationEnvelope(OperationEnvelope.SCHEMA, "copy-source-write", "native-copy-source", 1, 0,
                List.of(new OperationMutation("cell.set", "sheet-1", params)), Instant.now());
        operations.commit("native-copy-source", write, "copy-owner");

        var copied = catalog.copy("native-copy-source", new CopyWorkbookRequest("Copied", null, null), "copy-owner");

        assertNotEquals("native-copy-source", copied.unitId());
        var copiedArtifact = catalog.getArtifact(copied.unitId(), "copy-owner");
        assertEquals(0, copiedArtifact.getWorkbookRevision());
        try (var zip = new java.util.zip.ZipFile(copiedArtifact.getStoragePath())) {
            var documentFactory = javax.xml.parsers.DocumentBuilderFactory.newInstance();
            documentFactory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            documentFactory.setNamespaceAware(true);
            try (var xml = zip.getInputStream(zip.getEntry("xl/worksheets/sheet1.xml"))) {
                var document = documentFactory.newDocumentBuilder().parse(xml);
                var values = document.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "v");
                assertEquals(1, values.getLength());
                assertEquals(42, Double.parseDouble(values.item(0).getTextContent()));
            }
        }
        kernel.close();
        var opened = operations.open(copied.unitId(), "copy-owner");
        assertEquals(copied.unitId(), opened.manifest().path("unitId").asText());
        assertEquals("Copied", opened.manifest().path("name").asText());
        kernel.call("open", mapper.createObjectNode().set("manifest", opened.manifest()));
        var load = mapper.createObjectNode().put("unitId", copied.unitId()).put("revision", opened.revision());
        load.set("page", operations.page(copied.unitId(), opened.revision(), "sheet-1", 0, 0, "copy-owner"));
        kernel.call("page.load", load);
        var read = mapper.createObjectNode().put("unitId", copied.unitId()).put("revision", opened.revision());
        read.putObject("address").put("sheetId", "sheet-1").put("row", 0).put("column", 0);
        assertEquals(42, kernel.call("cell.get", read).path("cell").path("value").asInt());
        var source = operations.open("native-copy-source", "copy-owner");
        assertEquals("native-copy-source", source.manifest().path("unitId").asText());
        assertEquals("Source", source.manifest().path("name").asText());
        assertEquals(1, source.revision());
    }

    @Test
    void exportPersistsNativeBytesBoundToRevisionAndRejectsStaleDownload() throws Exception {
        String unitId = "native-export";
        var sheets = mapper.readTree("""
                [{"sheetId":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,"metadata":{}}]
                """);
        catalog.create(new CreateWorkbookRequest(unitId, "Export", sheets, null, null, null, null), "export-owner");

        var exported = catalog.exportArtifact(unitId, 0, "export.xlsx", "xlsx", "export-owner");

        var stored = catalog.getArtifact(unitId, "export-owner");
        byte[] bytes = java.nio.file.Files.readAllBytes(java.nio.file.Path.of(stored.getStoragePath()));
        assertEquals(0, stored.getWorkbookRevision());
        assertEquals(bytes.length, exported.byteLength());
        assertEquals(java.util.HexFormat.of().formatHex(java.security.MessageDigest.getInstance("SHA-256").digest(bytes)), exported.checksum());
        try (var zip = new java.util.zip.ZipFile(stored.getStoragePath())) {
            assertNotNull(zip.getEntry("[Content_Types].xml"));
            assertNotNull(zip.getEntry("xl/workbook.xml"));
        }
        var params = mapper.createObjectNode().put("sheetId", "sheet-1").put("row", 0).put("column", 0);
        params.putObject("value").put("value", 42);
        operations.commit(unitId, new OperationEnvelope(OperationEnvelope.SCHEMA, "export-next-revision", unitId, 1, 0,
                List.of(new OperationMutation("cell.set", "sheet-1", params)), Instant.now()), "export-owner");
        assertEquals("CONFLICT", assertThrows(ServiceException.class,
                () -> catalog.getArtifact(unitId, "export-owner")).code());
        assertEquals("CONFLICT", assertThrows(ServiceException.class,
                () -> catalog.exportArtifact(unitId, 0, "stale.xlsx", "xlsx", "export-owner")).code());
        var refreshed = catalog.exportArtifact(unitId, 1, null, null, "export-owner");
        assertEquals("export.xlsx", refreshed.fileName());
        assertEquals(1, refreshed.revision());
        assertEquals("xlsx", refreshed.nativeMetadata().path("format").asText());
    }

    @Test
    void repeatedExportDeletesReplacedArtifactAfterCommit() throws Exception {
        String unitId = "native-export-repeat";
        catalog.create(new CreateWorkbookRequest(unitId, "Export", null, null, null, null, null), "export-owner");

        var first = catalog.exportArtifact(unitId, 0, "first.xlsx", "xlsx", "export-owner");
        java.nio.file.Path firstPath = java.nio.file.Path.of(catalog.getArtifact(unitId, "export-owner").getStoragePath());
        assertTrue(java.nio.file.Files.isRegularFile(firstPath));

        var second = catalog.exportArtifact(unitId, 0, "second.xlsx", "xlsx", "export-owner");
        java.nio.file.Path secondPath = java.nio.file.Path.of(catalog.getArtifact(unitId, "export-owner").getStoragePath());
        assertNotEquals(firstPath, secondPath);
        assertEquals(first.checksum(), second.checksum());
        assertFalse(java.nio.file.Files.exists(firstPath));
        assertTrue(java.nio.file.Files.isRegularFile(secondPath));
    }

    @Test
    void exportRollbackRetainsPreviousArtifactAndRemovesCandidate() throws Exception {
        String unitId = "native-export-rollback";
        catalog.create(new CreateWorkbookRequest(unitId, "Export", null, null, null, null, null), "export-owner");
        catalog.exportArtifact(unitId, 0, "first.xlsx", "xlsx", "export-owner");
        java.nio.file.Path firstPath = java.nio.file.Path.of(catalog.getArtifact(unitId, "export-owner").getStoragePath());
        AtomicReference<java.nio.file.Path> candidatePath = new AtomicReference<>();

        new TransactionTemplate(transactionManager).executeWithoutResult(status -> {
            catalog.exportArtifact(unitId, 0, "second.xlsx", "xlsx", "export-owner");
            candidatePath.set(java.nio.file.Path.of(catalog.getArtifact(unitId, "export-owner").getStoragePath()));
            assertTrue(java.nio.file.Files.isRegularFile(candidatePath.get()));
            status.setRollbackOnly();
        });

        assertTrue(java.nio.file.Files.isRegularFile(firstPath));
        assertFalse(java.nio.file.Files.exists(candidatePath.get()));
        assertEquals(firstPath.toString(), catalog.getArtifact(unitId, "export-owner").getStoragePath());
    }

    @Test
    void purgeDeletesCapturedArtifactAfterCommit() throws Exception {
        String unitId = "native-export-purge";
        catalog.create(new CreateWorkbookRequest(unitId, "Export", null, null, null, null, null), "export-owner");
        catalog.exportArtifact(unitId, 0, "purge.xlsx", "xlsx", "export-owner");
        java.nio.file.Path artifactPath = java.nio.file.Path.of(catalog.getArtifact(unitId, "export-owner").getStoragePath());
        assertTrue(java.nio.file.Files.isRegularFile(artifactPath));

        catalog.moveToTrash(unitId, "export-owner");
        catalog.purge(unitId, "export-owner");

        assertFalse(java.nio.file.Files.exists(artifactPath));
        assertTrue(artifacts.findById(unitId).isEmpty());
    }

    @Test
    void purgeRollbackRetainsCapturedArtifact() throws Exception {
        String unitId = "native-export-purge-rollback";
        catalog.create(new CreateWorkbookRequest(unitId, "Export", null, null, null, null, null), "export-owner");
        catalog.exportArtifact(unitId, 0, "purge.xlsx", "xlsx", "export-owner");
        java.nio.file.Path artifactPath = java.nio.file.Path.of(catalog.getArtifact(unitId, "export-owner").getStoragePath());
        catalog.moveToTrash(unitId, "export-owner");

        new TransactionTemplate(transactionManager).executeWithoutResult(status -> {
            catalog.purge(unitId, "export-owner");
            status.setRollbackOnly();
        });

        assertTrue(java.nio.file.Files.isRegularFile(artifactPath));
        assertEquals(artifactPath.toString(), catalog.getArtifact(unitId, "export-owner").getStoragePath());
    }

    @Test
    void viewerCannotPublishNativeArtifact() {
        String unitId = "native-export-role";
        catalog.create(new CreateWorkbookRequest(unitId, "Export", null, null, null, null, null), "export-owner");
        var share = shares.create(unitId, new ShareCreateRequest("viewer", Instant.now().plusSeconds(600)), "export-owner");
        var error = assertThrows(ServiceException.class, () -> catalog.exportArtifact(
                unitId, 0, "forbidden.xlsx", "xlsx", "guest:" + share.shareId()));
        assertEquals("FORBIDDEN", error.code());
        assertEquals("NOT_FOUND", assertThrows(ServiceException.class,
                () -> catalog.getArtifact(unitId, "export-owner")).code());
    }
}
