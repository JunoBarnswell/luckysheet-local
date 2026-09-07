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
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.context.TestPropertySource;

import java.time.Instant;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@TestPropertySource(properties = "DATABASE_URL=jdbc:h2:mem:native_catalog;DB_CLOSE_DELAY=-1")
class WorkbookCatalogServiceTest extends NativeKernelIntegrationTestSupport {
    @Autowired private WorkbookCatalogService catalog;
    @Autowired private WorkbookOperationService operations;
    @Autowired private WorkbookEntityRepository workbooks;
    @Autowired private WorkbookManifestEntityRepository manifests;
    @Autowired private KernelHostClient kernel;
    @Autowired private GuestShareService shares;
    @Autowired private ObjectMapper mapper;

    @Test
    void createPublishesNativeManifestAndReopensAfterProcessRestart() {
        var created = catalog.create(new CreateWorkbookRequest("native-catalog", "Catalog", null, null, null, null), "owner");
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
        var request = new CreateWorkbookRequest("invalid-native-catalog", "Invalid", sheets, null, null, null);
        var error = assertThrows(KernelHostException.class, () -> catalog.create(request, "owner"));
        assertEquals("SHEET_INVALID", error.code());
        assertFalse(workbooks.existsById(request.unitId()));
        assertTrue(manifests.findByUnitIdAndRevision(request.unitId(), 0).isEmpty());
    }

    @Test
    void nativeCopyHasIndependentIdentityAndCellsAcrossPersistedReopen() throws Exception {
        var sheets = mapper.readTree("""
                [{"sheetId":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,"metadata":{}}]
                """);
        catalog.create(new CreateWorkbookRequest("native-copy-source", "Source", sheets, null, null, null), "copy-owner");
        var params = mapper.createObjectNode().put("sheetId", "sheet-1").put("row", 0).put("column", 0);
        params.putObject("value").put("value", 42);
        var write = new OperationEnvelope(OperationEnvelope.SCHEMA, "copy-source-write", "native-copy-source", 1, 0,
                List.of(new OperationMutation("cell.set", "sheet-1", params)), Instant.now());
        operations.commit("native-copy-source", write, "copy-owner");

        var copied = catalog.copy("native-copy-source", new CopyWorkbookRequest("Copied", null, null), "copy-owner");

        assertNotEquals("native-copy-source", copied.unitId());
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
        catalog.create(new CreateWorkbookRequest(unitId, "Export", sheets, null, null, null), "export-owner");

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
    }

    @Test
    void viewerCannotPublishNativeArtifact() {
        String unitId = "native-export-role";
        catalog.create(new CreateWorkbookRequest(unitId, "Export", null, null, null, null), "export-owner");
        var share = shares.create(unitId, new ShareCreateRequest("viewer", Instant.now().plusSeconds(600)), "export-owner");
        var error = assertThrows(ServiceException.class, () -> catalog.exportArtifact(
                unitId, 0, "forbidden.xlsx", "xlsx", "guest:" + share.shareId()));
        assertEquals("FORBIDDEN", error.code());
        assertEquals("NOT_FOUND", assertThrows(ServiceException.class,
                () -> catalog.getArtifact(unitId, "export-owner")).code());
    }
}
