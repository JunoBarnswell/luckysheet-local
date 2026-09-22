package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.CopyWorkbookRequest;
import com.xc.luckysheet.server.contract.OperationOrigin;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RestoreRequest;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.ShareCreateRequest;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.DataBlockRow;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import com.xc.luckysheet.server.store.WorkbookStore;
import com.xc.luckysheet.server.store.OperationRow;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.TestPropertySource;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Instant;
import java.util.HexFormat;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@SpringBootTest
@TestPropertySource(properties = {
        "DATABASE_URL=jdbc:h2:mem:block_authorization;DB_CLOSE_DELAY=-1",
        "DATABASE_USERNAME=sa",
        "DATABASE_PASSWORD=",
        "JPA_DDL_AUTO=validate",
        "FLYWAY_BASELINE_ON_MIGRATE=false",
        "luckysheet.auth.mode=oidc",
        "AUTH_ISSUER=https://issuer.test",
        "AUTH_AUDIENCE=test",
        "AUTH_JWKS_URL=https://issuer.test/.well-known/jwks.json",
        "COORDINATION_MULTI_INSTANCE=false",
        "COORDINATION_REDIS_ENABLED=false"
})
class WorkbookDataBlockAuthorizationIntegrationTest {
    @Autowired private WorkbookCatalogService catalog;
    @Autowired private GuestShareService shares;
    @Autowired private WorkbookDataBlockCommitService commits;
    @Autowired private WorkbookDataBlockStore blocks;
    @Autowired private ObjectMapper mapper;
    @Autowired private WorkbookStore workbooks;
    @Autowired private WorkbookOperationService operations;

    @Test
    void revokedEditorCannotCommitBytesReadBeforeTheWriteBoundary() throws Exception {
        String unitId = "block-revoked";
        String owner = "owner-revoked";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), owner);
        var share = shares.create(unitId, new ShareCreateRequest("editor", Instant.now().plusSeconds(600)), owner);
        byte[] content = "blocked-after-revocation".getBytes(StandardCharsets.UTF_8);

        shares.revoke(unitId, share.shareId(), owner);
        ServiceException error = assertThrows(ServiceException.class, () -> commits.commit(
                row(unitId, content), WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_BYTES,
                WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_COUNT, "guest:" + share.shareId()));

        assertEquals("FORBIDDEN", error.code());
        assertTrue(blocks.find(unitId, "source", "block").isEmpty());
    }

    @Test
    void editorCommitThatCompletesBeforeRevocationRemainsPersisted() throws Exception {
        String unitId = "block-committed";
        String owner = "owner-committed";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), owner);
        var share = shares.create(unitId, new ShareCreateRequest("editor", Instant.now().plusSeconds(600)), owner);
        byte[] content = "committed-before-revocation".getBytes(StandardCharsets.UTF_8);

        var metadata = commits.commit(row(unitId, content), WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_BYTES,
                WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_COUNT, "guest:" + share.shareId());
        shares.revoke(unitId, share.shareId(), owner);

        assertEquals(content.length, metadata.byteLength());
        assertEquals(HexFormat.of().formatHex(content), HexFormat.of().formatHex(blocks.find(unitId, "source", "block").orElseThrow().content()));
    }

    @Test
    void sameIdentityIsIdempotentButRejectsDifferentBytes() throws Exception {
        String unitId = "block-immutable";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        DataBlockRow original = row(unitId, "original".getBytes(StandardCharsets.UTF_8));
        var first = commits.commit(original, 100, 10, "owner");
        var repeated = commits.commit(original, 100, 10, "owner");
        assertEquals(first.checksum(), repeated.checksum());
        assertEquals(first.byteLength(), repeated.byteLength());

        ServiceException error = assertThrows(ServiceException.class, () -> commits.commit(
                row(unitId, "modified".getBytes(StandardCharsets.UTF_8)), 100, 10, "owner"));
        assertEquals("DATA_BLOCK_IMMUTABLE", error.code());
        assertArrayEquals(original.content(), blocks.find(unitId, "source", "block").orElseThrow().content());
    }

    @Test
    void unreferencedStagingBlockCanBeDeletedIdempotently() throws Exception {
        String unitId = "block-delete-staging";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        commits.commit(row(unitId, new byte[] {1}), 100, 10, "owner");
        commits.delete(unitId, "source", "block", "owner");
        commits.delete(unitId, "source", "block", "owner");
        assertTrue(blocks.find(unitId, "source", "block").isEmpty());
    }

    @Test
    void currentSnapshotReferencePreventsDeletion() throws Exception {
        String unitId = "block-current-reference";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        DataBlockRow block = row(unitId, new byte[] {1});
        commits.commit(block, 100, 10, "owner");
        workbooks.updateWorkbook(unitId, 1, snapshotWithBlock(unitId, block).toString(), 1, Instant.now());
        assertReferenced(block);
    }

    @Test
    void historicalCheckpointReferencePreventsDeletionAfterSourceRemoval() throws Exception {
        String unitId = "block-checkpoint-reference";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        DataBlockRow block = row(unitId, new byte[] {1});
        commits.commit(block, 100, 10, "owner");
        String historical = snapshotWithBlock(unitId, block).toString();
        workbooks.insertCheckpoint(unitId, 1, historical, sha256(historical), Instant.now());
        workbooks.updateWorkbook(unitId, 2, snapshot(unitId).toString(), 2, Instant.now());
        assertReferenced(block);
    }

    @Test
    void retainedOperationReferencePreventsDeletionAfterSourceRemoval() throws Exception {
        String unitId = "block-operation-reference";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        DataBlockRow block = row(unitId, new byte[] {1});
        commits.commit(block, 100, 10, "owner");
        ObjectNode params = mapper.createObjectNode();
        params.set("source", snapshotWithBlock(unitId, block).path("dataModel").path("sources").get(0));
        Instant now = Instant.now();
        var envelope = new CommittedOperationEnvelope("session", "OperationEnvelope", "op-" + unitId,
                unitId, "owner", OperationOrigin.CLIENT, 1, 0, 1,
                List.of(new CommittedOperationMutation("dataSource.add", "sheet-1", params, List.of())), now, now);
        workbooks.insertOperation(new OperationRow(envelope.operationId(), unitId, 1, "owner", "session", 1, 0,
                mapper.writeValueAsString(envelope), now));
        workbooks.updateWorkbook(unitId, 2, snapshot(unitId).toString(), 2, now);
        assertReferenced(block);
    }

    @Test
    void corruptRetainedHistoryRejectsDeletionWithoutLosingBytes() throws Exception {
        String unitId = "block-corrupt-history";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        DataBlockRow block = row(unitId, new byte[] {1});
        commits.commit(block, 100, 10, "owner");
        workbooks.insertCheckpoint(unitId, 1, "{", sha256("{"), Instant.now());
        ServiceException error = assertThrows(ServiceException.class, () -> commits.delete(unitId, "source", "block", "owner"));
        assertEquals("DATA_BLOCK_REFERENCE_CHECK_FAILED", error.code());
        assertArrayEquals(block.content(), blocks.find(unitId, "source", "block").orElseThrow().content());
    }

    @Test
    void referenceMatchingDoesNotCombineFieldsAcrossObjectsOrMatchTextSubstrings() throws Exception {
        String unitId = "block-exact-reference";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        commits.commit(row(unitId, new byte[] {1}), 100, 10, "owner");
        String history = "{\"items\":[{\"id\":\"block-other\",\"dataSourceId\":\"source\"},"
                + "{\"id\":\"block\",\"nested\":{\"dataSourceId\":\"source\"}}],"
                + "\"text\":\"{\\\"id\\\":\\\"block\\\",\\\"dataSourceId\\\":\\\"source\\\"}\"}";
        workbooks.insertCheckpoint(unitId, 1, history, sha256(history), Instant.now());
        commits.delete(unitId, "source", "block", "owner");
        assertTrue(blocks.find(unitId, "source", "block").isEmpty());
    }

    @Test
    void committedSourceReferencesOnlyUploadedMetadataAndProtectsItsBytes() throws Exception {
        String unitId = "block-publish-success";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        DataBlockRow block = row(unitId, new byte[] {1});
        commits.commit(block, 100, 10, "owner");
        OperationEnvelope operation = sourceOperation(unitId, block);
        operations.commit(unitId, operation, "owner");
        assertEquals(1, workbooks.find(unitId).orElseThrow().revision());
        assertEquals("source", operations.readSnapshot(unitId, "owner").snapshot().path("dataModel").path("sources").get(0).path("id").asText());
        assertReferenced(block);
    }

    @Test
    void deletingStagingBytesBeforePublicationRejectsTheOperationAndAllowsReupload() throws Exception {
        String unitId = "block-publish-after-delete";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        DataBlockRow block = row(unitId, new byte[] {1});
        commits.commit(block, 100, 10, "owner");
        OperationEnvelope operation = sourceOperation(unitId, block);
        commits.delete(unitId, "source", "block", "owner");
        ServiceException error = assertThrows(ServiceException.class, () -> operations.commit(unitId, operation, "owner"));
        assertEquals("DATA_BLOCK_MISSING", error.code());
        assertEquals(0, workbooks.find(unitId).orElseThrow().revision());
        assertTrue(workbooks.findOperation(operation.operationId()).isEmpty());
        commits.commit(block, 100, 10, "owner");
        operations.commit(unitId, operation, "owner");
        assertEquals(1, workbooks.find(unitId).orElseThrow().revision());
    }

    @Test
    void mismatchedDescriptorsAndBlocksFromAnotherWorkbookCannotBePublished() throws Exception {
        String unitId = "block-publish-mismatch";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        DataBlockRow block = row(unitId, new byte[] {1});
        commits.commit(block, 100, 10, "owner");
        for (String property : List.of("checksum", "byteLength")) {
            OperationEnvelope operation = sourceOperation(unitId, block);
            ObjectNode descriptor = (ObjectNode) operation.mutations().get(0).params().path("source").path("blocks").get(0);
            if (property.equals("checksum")) descriptor.put(property, "f".repeat(64));
            else descriptor.put(property, 2);
            ServiceException error = assertThrows(ServiceException.class, () -> operations.commit(unitId, operation, "owner"));
            assertEquals("DATA_BLOCK_METADATA_MISMATCH", error.code());
        }
        assertEquals(0, workbooks.find(unitId).orElseThrow().revision());
        String otherUnit = "block-publish-other-workbook";
        catalog.create(new CreateWorkbookRequest(otherUnit, "Blocks", snapshot(otherUnit)), "owner");
        OperationEnvelope foreignReference = sourceOperation(otherUnit, row(otherUnit, block.content()));
        ServiceException missing = assertThrows(ServiceException.class, () -> operations.commit(otherUnit, foreignReference, "owner"));
        assertEquals("DATA_BLOCK_MISSING", missing.code());
        assertEquals(0, workbooks.find(otherUnit).orElseThrow().revision());
        assertArrayEquals(block.content(), blocks.find(unitId, "source", "block").orElseThrow().content());
    }

    @Test
    void initialSnapshotCannotPublishReferencesBeforeTheirBytesExist() throws Exception {
        String unitId = "block-create-missing";
        DataBlockRow block = row(unitId, new byte[] {1});
        ServiceException error = assertThrows(ServiceException.class,
                () -> catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshotWithBlock(unitId, block)), "owner"));
        assertEquals("DATA_BLOCK_MISSING", error.code());
        assertTrue(workbooks.find(unitId).isEmpty());
    }

    @Test
    void workbookCopyCopiesPublishedBlocksWithoutSharingTheirIdentity() throws Exception {
        String sourceUnitId = "block-copy-source";
        catalog.create(new CreateWorkbookRequest(sourceUnitId, "Blocks", snapshot(sourceUnitId)), "owner");
        DataBlockRow sourceBlock = row(sourceUnitId, new byte[] {1, 2, 3});
        commits.commit(sourceBlock, 100, 10, "owner");
        operations.commit(sourceUnitId, sourceOperation(sourceUnitId, sourceBlock), "owner");

        var copied = catalog.copy(sourceUnitId, new CopyWorkbookRequest("Blocks Copy", null, null), "owner");
        DataBlockRow copiedBlock = blocks.find(copied.unitId(), sourceBlock.sourceId(), sourceBlock.blockId()).orElseThrow();

        assertEquals(sourceBlock.checksum(), copiedBlock.checksum());
        assertEquals(sourceBlock.byteLength(), copiedBlock.byteLength());
        assertArrayEquals(sourceBlock.content(), copiedBlock.content());
        assertReferenced(new DataBlockRow(copied.unitId(), copiedBlock.sourceId(), copiedBlock.blockId(), copiedBlock.checksum(),
                copiedBlock.byteLength(), copiedBlock.content(), copiedBlock.createdAt(), copiedBlock.updatedAt()));
    }

    @Test
    void restoreRejectsRetainedReferencesWhoseBytesAreMissing() throws Exception {
        String unitId = "block-restore-missing";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        String historical = snapshotWithBlock(unitId, row(unitId, new byte[] {1})).toString();
        workbooks.insertCheckpoint(unitId, 1, historical, sha256(historical), Instant.now());
        workbooks.updateWorkbook(unitId, 2, snapshot(unitId).toString(), 2, Instant.now());
        ServiceException error = assertThrows(ServiceException.class,
                () -> operations.restore(unitId, new RestoreRequest(1, "Restore historical blocks"), "owner"));
        assertEquals("DATA_BLOCK_MISSING", error.code());
        assertEquals(2, workbooks.find(unitId).orElseThrow().revision());
        assertTrue(workbooks.listOperations(unitId).isEmpty());
    }

    private OperationEnvelope sourceOperation(String unitId, DataBlockRow block) throws Exception {
        ObjectNode params = mapper.createObjectNode();
        params.set("source", snapshotWithBlock(unitId, block).path("dataModel").path("sources").get(0));
        return new OperationEnvelope("block-publish-session", OperationEnvelope.SCHEMA, "publish-" + unitId,
                unitId, 1, 0, List.of(new OperationMutation("dataSource.add", "sheet-1", params)), Instant.now());
    }

    private void assertReferenced(DataBlockRow block) {
        ServiceException error = assertThrows(ServiceException.class,
                () -> commits.delete(block.unitId(), block.sourceId(), block.blockId(), "owner"));
        assertEquals("DATA_BLOCK_REFERENCED", error.code());
        assertArrayEquals(block.content(), blocks.find(block.unitId(), block.sourceId(), block.blockId()).orElseThrow().content());
    }

    private String sha256(String document) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(document.getBytes(StandardCharsets.UTF_8)));
    }

    private ObjectNode snapshotWithBlock(String unitId, DataBlockRow block) throws Exception {
        ObjectNode root = (ObjectNode) snapshot(unitId);
        ObjectNode source = ((ArrayNode) root.path("dataModel").path("sources")).addObject();
        source.put("schema", "DataSourceManifest").put("version", 1).put("id", block.sourceId())
                .put("name", "Source").put("kind", "chunked-table").put("rowCount", 1).put("blockRowCount", 65536).put("revision", 0);
        source.putArray("fields").addObject().put("id", "value").put("name", "Value").put("ordinal", 0).put("type", "mixed");
        // The order deliberately differs from frontend JSON serialization.
        source.putArray("blocks").addObject().put("dataSourceId", block.sourceId()).put("id", block.blockId())
                .put("startRow", 0).put("rowCount", 1).put("storageKey", block.blockId()).put("checksum", block.checksum())
                .put("byteLength", block.byteLength()).put("encoding", "columnar-v1").put("revision", 0);
        return root;
    }

    private DataBlockRow row(String unitId, byte[] content) throws Exception {
        return new DataBlockRow(unitId, "source", "block", HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content)),
                content.length, content, Instant.now(), Instant.now());
    }

    private com.fasterxml.jackson.databind.JsonNode snapshot(String unitId) throws Exception {
        return mapper.readTree("{\"schema\":\"WorkbookSnapshot\",\"version\":9,\"unitId\":\"" + unitId
                + "\",\"name\":\"Blocks\",\"dimensionMetrics\":{\"normalFontFamily\":\"Calibri\",\"normalFontSizePx\":14.6666666667,\"maximumDigitWidthPx\":7},\"calculationSettings\":{},\"editingOptions\":{\"allowEditDirectly\":true,\"moveAfterEnter\":true,\"enterDirection\":\"down\",\"formulaAutoComplete\":true,\"valueAutoComplete\":true,\"fixedDecimalPlaces\":null},\"dataModel\":{\"sources\":[],\"tables\":[],\"relationships\":[],\"views\":[]},\"sheets\":[{\"kind\":\"worksheet\",\"id\":\"sheet-1\",\"name\":\"Sheet1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{},\"merges\":[],\"pane\":{\"kind\":\"none\"},\"defaultRowHeightPx\":20,\"defaultColumnWidthPx\":64,\"pivots\":[],\"sparklines\":[],\"drawings\":[],\"drawingPayloads\":{},\"review\":{\"notesByCell\":{},\"notesById\":{},\"threadIdsByCell\":{},\"threadsById\":{}}}]}");
    }
}
