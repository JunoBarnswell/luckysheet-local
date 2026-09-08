package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.NativeKernelIntegrationTestSupport;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationIntent;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.ShareCreateRequest;
import com.xc.luckysheet.server.persistence.WorkbookManifestEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookPageEntityRepository;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.context.TestPropertySource;

import java.time.Instant;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@TestPropertySource(properties = "DATABASE_URL=jdbc:h2:mem:native_operations;DB_CLOSE_DELAY=-1")
class WorkbookOperationServiceTest extends NativeKernelIntegrationTestSupport {
    @Autowired private WorkbookCatalogService catalog;
    @Autowired private WorkbookOperationService operations;
    @Autowired private WorkbookStore store;
    @Autowired private WorkbookManifestEntityRepository manifests;
    @Autowired private WorkbookPageEntityRepository pages;
    @Autowired private GuestShareService shares;
    @Autowired private KernelHostClient kernel;
    @Autowired private ObjectMapper mapper;

    @Test
    void nativeCommitPersistsPagesAndServerMetadataAndReopensAfterProcessRestart() throws Exception {
        String unitId = "native-commit";
        create(unitId);
        assertEquals(1, operations.accessProjection(unitId, "owner").nextClientSequence());
        var submitted = operation(unitId, "native-commit-op", 1, 0, set(1, 2, 42));
        Instant before = Instant.now();
        var result = operations.commit(unitId, submitted, "owner");
        assertTrue(result.committed());
        assertEquals("owner", result.operation().actorId());
        assertEquals(result.operation().committedAt(), result.operation().createdAt());
        assertFalse(result.operation().committedAt().isBefore(before));
        var range = result.operation().mutations().getFirst().affectedRanges().getFirst();
        assertEquals(1, range.startRow());
        assertEquals(2, range.startColumn());
        assertEquals(1, store.find(unitId).orElseThrow().revision());
        assertTrue(store.findOperation(submitted.operationId()).isPresent());
        assertTrue(manifests.findByUnitIdAndRevision(unitId, 1).isPresent());
        assertEquals(2, operations.accessProjection(unitId, "owner").nextClientSequence());

        kernel.close();
        var reopened = operations.open(unitId, "owner");
        assertEquals(1, reopened.revision());
        assertEquals(11, reopened.manifest().path("version").asInt());
        reopenNative(unitId, "owner", true);
        assertEquals(42, cell(unitId, 1, 1, 2).path("value").asInt());
        var duplicate = operations.commit(unitId, submitted, "owner");
        assertFalse(duplicate.committed());
        assertEquals(1, store.find(unitId).orElseThrow().revision());
    }

    @Test
    void restoreReusesHistoricalPagesInANewRevisionAfterRestart() throws Exception {
        String unitId = "native-restore-pages";
        create(unitId);
        operations.commit(unitId, operation(unitId, "restore-before", 1, 0, set(0, 0, 42)), "owner");
        operations.commit(unitId, operation(unitId, "restore-after", 2, 1, set(0, 0, 99)), "owner");
        kernel.close();
        operations.restore(unitId, new com.xc.luckysheet.server.contract.RestoreRequest(1, "Recover previous content"), "owner");
        assertEquals(3, store.find(unitId).orElseThrow().revision());
        kernel.close();
        reopenNative(unitId, "owner", true);
        assertEquals(42, cell(unitId, 3, 0, 0).path("value").asInt());
        assertEquals(2, operations.readRevision(unitId, 2, "owner").revision());
        assertThrows(ServiceException.class, () -> operations.restore(unitId,
                new com.xc.luckysheet.server.contract.RestoreRequest(4, "Invalid future"), "owner"));
        assertEquals(3, store.find(unitId).orElseThrow().revision());
    }

    @Test
    void commenterCannotCommitEditorMutationAndRejectionIsAudited() throws Exception {
        String unitId = "native-role";
        create(unitId);
        var share = shares.create(unitId, new ShareCreateRequest("commenter", Instant.now().plusSeconds(600)), "owner");
        String actor = "guest:" + share.shareId();
        var request = operation(unitId, "native-role-op", 1, 0, set(0, 0, 12));
        var error = assertThrows(ServiceException.class, () -> operations.commit(unitId, request, actor));
        assertEquals("FORBIDDEN", error.code());
        assertEquals(0, store.find(unitId).orElseThrow().revision());
        assertTrue(store.findOperation(request.operationId()).isEmpty());
        assertTrue(store.listAudit(unitId, 20).stream().anyMatch(audit ->
                request.operationId().equals(audit.operationId()) && "REJECTED".equals(audit.outcome())));
    }

    @Test
    void serverRenameUsesCanonicalRevisionAndTheNextActorSequence() throws Exception {
        String unitId = "native-server-rename";
        create(unitId);
        operations.commit(unitId, operation(unitId, "rename-prior-edit", 1, 0, set(0, 0, 42)), "owner");
        var params = mapper.createObjectNode().put("name", "Renamed once");

        var result = operations.commitServerMutation(unitId,
                new OperationMutation("workbook.renamed", "workbook", params), "owner", "workbook-rename");

        assertTrue(result.committed());
        assertEquals(2, result.operation().clientSequence());
        assertEquals(2, result.operation().revision());
        assertEquals("Renamed once", result.changeSet().path("manifest").path("name").asText());
        assertEquals("Renamed once", store.find(unitId).orElseThrow().name());
        kernel.close();
        var reopened = operations.open(unitId, "owner");
        assertEquals(2, reopened.revision());
        assertEquals("Renamed once", reopened.manifest().path("name").asText());
    }

    @Test
    void viewerCannotUseServerRenameAndNoRevisionIsPublished() throws Exception {
        String unitId = "native-server-rename-role";
        create(unitId);
        var share = shares.create(unitId, new ShareCreateRequest("viewer", Instant.now().plusSeconds(600)), "owner");
        String actor = "guest:" + share.shareId();
        var params = mapper.createObjectNode().put("name", "Forbidden rename");

        var error = assertThrows(ServiceException.class, () -> operations.commitServerMutation(unitId,
                new OperationMutation("workbook.renamed", "workbook", params), actor, "workbook-rename"));

        assertEquals("FORBIDDEN", error.code());
        assertEquals(0, store.find(unitId).orElseThrow().revision());
        assertTrue(manifests.findByUnitIdAndRevision(unitId, 1).isEmpty());
        assertEquals("Book", operations.open(unitId, "owner").manifest().path("name").asText());
    }

    @Test
    void invalidNativeMutationTailRejectsWholeBatchWithoutPersistentOrNativePartialWrite() throws Exception {
        String unitId = "native-invalid-batch";
        create(unitId);
        var request = operation(unitId, "native-invalid-op", 1, 0, set(0, 0, 12),
                new OperationMutation("not.a.command", "sheet-1", mapper.createObjectNode()));
        var error = assertThrows(KernelHostException.class, () -> operations.commit(unitId, request, "owner"));
        assertEquals("COMMAND_UNKNOWN", error.code());
        assertEquals(0, store.find(unitId).orElseThrow().revision());
        assertTrue(store.findOperation(request.operationId()).isEmpty());
        assertTrue(manifests.findByUnitIdAndRevision(unitId, 1).isEmpty());
        reopenNative(unitId, "owner", false);
        assertTrue(cell(unitId, 0, 0, 0).isNull());
    }

    @Test
    void staleRevisionCannotOverwriteCommittedCell() throws Exception {
        String unitId = "native-stale";
        create(unitId);
        operations.commit(unitId, operation(unitId, "native-first", 1, 0, set(0, 0, 42)), "owner");
        var stale = operation(unitId, "native-stale-op", 2, 0, set(0, 0, 99));
        var error = assertThrows(ServiceException.class, () -> operations.commit(unitId, stale, "owner"));
        assertEquals("CONFLICT", error.code());
        assertTrue(store.findOperation(stale.operationId()).isEmpty());
        reopenNative(unitId, "owner", true);
        assertEquals(42, cell(unitId, 1, 0, 0).path("value").asInt());
    }

    @Test
    void undoUsesPersistedNativeHistoryWithoutClientInverseMutations() throws Exception {
        String unitId = "native-undo";
        create(unitId);
        operations.commit(unitId, operation(unitId, "native-undo-target", 1, 0, set(0, 0, 42)), "owner");
        kernel.close();
        var undo = new OperationEnvelope(OperationEnvelope.SCHEMA, "native-undo-op", unitId, 2, 1,
                List.of(), Instant.now(), new OperationIntent(OperationIntent.UNDO, "native-undo-target", 0));

        var result = operations.commit(unitId, undo, "owner");

        assertTrue(result.committed());
        assertEquals(2, store.find(unitId).orElseThrow().revision());
        reopenNative(unitId, "owner", false);
        assertTrue(cell(unitId, 2, 0, 0).isNull());
    }

    @Test
    void undoRejectsOverlappingLaterPageWithoutPartialCommit() throws Exception {
        String unitId = "native-undo-conflict";
        create(unitId);
        operations.commit(unitId, operation(unitId, "undo-conflict-target", 1, 0, set(0, 0, 42)), "owner");
        operations.commit(unitId, operation(unitId, "undo-conflict-later", 2, 1, set(0, 0, 99)), "owner");
        var undo = new OperationEnvelope(OperationEnvelope.SCHEMA, "undo-conflict-request", unitId, 3, 2,
                List.of(), Instant.now(), new OperationIntent(OperationIntent.UNDO, "undo-conflict-target", 0));

        var error = assertThrows(KernelHostException.class, () -> operations.commit(unitId, undo, "owner"));

        assertEquals("UNDO_CONFLICT", error.code());
        assertEquals(2, store.find(unitId).orElseThrow().revision());
        assertTrue(store.findOperation(undo.operationId()).isEmpty());
        assertTrue(manifests.findByUnitIdAndRevision(unitId, 3).isEmpty());
        reopenNative(unitId, "owner", true);
        assertEquals(99, cell(unitId, 2, 0, 0).path("value").asInt());
    }

    @Test
    void undoCannotUseAnotherActorsCommittedHistory() throws Exception {
        String unitId = "native-undo-owner";
        create(unitId);
        operations.commit(unitId, operation(unitId, "native-owned-history", 1, 0, set(0, 0, 42)), "owner");
        var share = shares.create(unitId, new ShareCreateRequest("editor", Instant.now().plusSeconds(600)), "owner");
        var undo = new OperationEnvelope(OperationEnvelope.SCHEMA, "native-foreign-undo", unitId, 1, 1,
                List.of(), Instant.now(), new OperationIntent(OperationIntent.UNDO, "native-owned-history", 0));

        var error = assertThrows(ServiceException.class, () -> operations.commit(unitId, undo, "guest:" + share.shareId()));

        assertEquals("FORBIDDEN", error.code());
        assertEquals(1, store.find(unitId).orElseThrow().revision());
        assertTrue(store.findOperation(undo.operationId()).isEmpty());
    }

    @Test
    void missingCommittedPageFailsClosedAtPageReadAndCommandBoundary() throws Exception {
        String unitId = "native-page-missing";
        create(unitId);
        operations.commit(unitId, operation(unitId, "native-page-first", 1, 0, set(0, 0, 42)), "owner");
        var manifest = operations.open(unitId, "owner").manifest();
        var page = pages.findByUnitIdAndChecksum(unitId, manifest.path("pages").get(0).path("checksum").asText()).orElseThrow();
        pages.deleteById(page.getPageId());
        kernel.close();

        var readError = assertThrows(KernelHostException.class,
                () -> operations.page(unitId, 1, "sheet-1", 0, 0, "owner"));
        var request = operation(unitId, "native-page-rejected", 2, 1, set(0, 0, 99));
        var commitError = assertThrows(KernelHostException.class, () -> operations.commit(unitId, request, "owner"));

        assertEquals("PAGE_MISSING", readError.code());
        assertEquals("PAGE_MISSING", commitError.code());
        assertEquals(1, store.find(unitId).orElseThrow().revision());
        assertTrue(store.findOperation(request.operationId()).isEmpty());
        assertTrue(manifests.findByUnitIdAndRevision(unitId, 2).isEmpty());
    }

    private void create(String unitId) throws Exception {
        JsonNode sheets = mapper.readTree("""
                [{"sheetId":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,"metadata":{}}]
                """);
        catalog.create(new CreateWorkbookRequest(unitId, "Book", sheets, null, null, null, null), "owner");
    }

    private OperationMutation set(int row, int column, int value) {
        var params = mapper.createObjectNode().put("sheetId", "sheet-1").put("row", row).put("column", column);
        params.putObject("value").put("value", value);
        return new OperationMutation("cell.set", "sheet-1", params);
    }

    private OperationEnvelope operation(String unitId, String id, long sequence, long revision, OperationMutation... mutations) {
        return new OperationEnvelope(OperationEnvelope.SCHEMA, id, unitId, sequence, revision, List.of(mutations),
                Instant.parse("2000-01-01T00:00:00Z"));
    }

    private JsonNode cell(String unitId, long revision, int row, int column) {
        var params = mapper.createObjectNode().put("unitId", unitId).put("revision", revision);
        params.putObject("address").put("sheetId", "sheet-1").put("row", row).put("column", column);
        return kernel.call("cell.get", params).path("cell");
    }

    private void reopenNative(String unitId, String actor, boolean hasPage) {
        var opened = operations.open(unitId, actor);
        kernel.call("open", mapper.createObjectNode().set("manifest", opened.manifest()));
        if (hasPage) {
            var load = mapper.createObjectNode().put("unitId", unitId).put("revision", opened.revision());
            load.set("page", operations.page(unitId, opened.revision(), "sheet-1", 0, 0, actor));
            kernel.call("page.load", load);
        }
    }
}
