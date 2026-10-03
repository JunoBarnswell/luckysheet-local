package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.AccessPrincipal;
import com.xc.luckysheet.server.contract.AccessPrincipalKind;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeAccessGrant;
import com.xc.luckysheet.server.contract.RangeAccessLevel;
import com.xc.luckysheet.server.contract.RangeAccessRegionRequest;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.ShareCreateRequest;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.TestPropertySource;

import java.time.Instant;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@SpringBootTest
@TestPropertySource(properties = {
        "DATABASE_URL=jdbc:h2:mem:linked_data_acceptance;DB_CLOSE_DELAY=-1",
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
class LinkedDataAcceptanceIntegrationTest {
    @Autowired private WorkbookCatalogService catalog;
    @Autowired private GuestShareService shares;
    @Autowired private RangeAccessService rangeAccess;
    @Autowired private WorkbookOperationService operations;
    @Autowired private ObjectMapper mapper;
    @Autowired private AccessControlService acl;
    @Autowired private com.xc.luckysheet.server.store.WorkbookStore store;
    @Autowired private org.springframework.transaction.PlatformTransactionManager transactionManager;

    @Test
    void externalInputsRequireSourceMembershipAndWithholdHiddenInputs() throws Exception {
        String owner = "linked-owner", sourceId = "linked-source", targetId = "linked-target";
        catalog.create(new CreateWorkbookRequest(sourceId, "Range access", snapshot(sourceId)), owner);
        ObjectNode target = (ObjectNode) snapshot(targetId);
        ((com.fasterxml.jackson.databind.node.ArrayNode) target.path("dataModel").path("externalLinks")).add(mapper.readTree("""
          {"id":"link-1","token":"Source.xlsx","sourceUnitId":"linked-source","sheets":[{"token":"Sales","sheetId":"sheet-1"}]}
          """));
        catalog.create(new CreateWorkbookRequest(targetId, "Range access", target), owner);
        var share = shares.create(targetId, new ShareCreateRequest("editor", Instant.now().plusSeconds(600)), owner);
        String guest = "guest:" + share.shareId();
        JsonNode denied = node(operations.readExternalCalculationGraph(targetId, guest, List.of()), sourceId);
        assertEquals("denied", denied.path("state").asText()); assertFalse(denied.has("snapshot"));
        JsonNode graph = operations.readExternalCalculationGraph(targetId, owner, List.of());
        JsonNode input = node(graph, sourceId);
        com.xc.luckysheet.server.contract.WorkbookSnapshotValidator.requireCanonical(input.path("snapshot"), sourceId);
        assertEquals(owner, graph.path("subject").asText());
        assertEquals(0, input.path("revision").asLong());
        assertEquals(11, input.path("snapshot").path("version").asInt());
        assertTrue(input.path("snapshot").path("dataModel").path("externalLinks").isArray());
        assertEquals("SECRET-CELL-VALUE", input.path("snapshot").path("sheets").get(0).path("cells").path("0").path("1").path("value").asText());
    }

    @Test
    void deleteThreeDEndpointPersistsDerivedReferenceFacts() throws Exception {
        String unit = "three-d-lifecycle", owner = "three-d-owner";
        ObjectNode root = (ObjectNode) snapshot(unit);
        var sheets = (com.fasterxml.jackson.databind.node.ArrayNode) root.path("sheets");
        for (int i = 2; i <= 4; i++) {
            ObjectNode copy = ((ObjectNode) sheets.get(0)).deepCopy();
            copy.put("id", "sheet-" + i).put("name", "Sheet" + i); copy.putObject("cells"); sheets.add(copy);
        }
        ObjectNode formula = mapper.createObjectNode().put("value", (String) null).put("formula", "=SUM(Sheet1:Sheet3!A1)");
        ((ObjectNode) sheets.get(3).path("cells")).putObject("0").set("0", formula);
        catalog.create(new CreateWorkbookRequest(unit, "Range access", root), owner);
        var operation = new OperationEnvelope("lifecycle-session", OperationEnvelope.SCHEMA, "delete-endpoint", unit, 1, 0,
            List.of(new OperationMutation("sheet.remove", "sheet-1", mapper.readTree("{\"id\":\"sheet-1\"}"))), Instant.now());
        JsonNode preimage = operations.readSnapshot(unit, owner).snapshot();
        var result = operations.commit(unit, operation, owner);
        JsonNode persisted = operations.readSnapshot(unit, owner).snapshot();
        assertEquals("=SUM(Sheet2:Sheet3!A1)", persisted.path("sheets").get(2).path("cells").path("0").path("0").path("formula").asText());
        assertEquals(1, operations.readSnapshot(unit, owner).revision());
        assertEquals(10, result.operation().mutations().getFirst().structuralPatch().version());
        assertEquals(1, result.operation().mutations().getFirst().structuralPatch().formulaOwnerDeltas().size());
        ObjectNode restoreParams = mapper.createObjectNode(); restoreParams.set("sheet", preimage.path("sheets").get(0)); restoreParams.put("index", 0);
        ObjectNode tampered = restoreParams.deepCopy(); ((ObjectNode) tampered.get("sheet")).put("name", "Tampered");
        var intent = new com.xc.luckysheet.server.contract.OperationIntent("undo", "delete-endpoint", 0);
        var invalidUndo = new OperationEnvelope("lifecycle-session", OperationEnvelope.SCHEMA, "tampered-undo", unit, 2, 1,
            List.of(new OperationMutation("sheet.restore", "sheet-1", tampered)), Instant.now(), intent);
        assertThrows(ServiceException.class, () -> operations.commit(unit, invalidUndo, owner));
        assertEquals(1, operations.readSnapshot(unit, owner).revision());
        var undo = new OperationEnvelope("lifecycle-session", OperationEnvelope.SCHEMA, "undo-endpoint", unit, 3, 1,
            List.of(new OperationMutation("sheet.restore", "sheet-1", restoreParams)), Instant.now(), intent);
        operations.commit(unit, undo, owner);
        assertEquals(preimage, operations.readSnapshot(unit, owner).snapshot());
        assertEquals(2, operations.readSnapshot(unit, owner).revision());
    }
    @Test
    void recordCalculationInputsRequireCompleteSourceRangeReadAccess() throws Exception {
        String owner = "record-input-owner", reader = "record-input-reader", sourceId = "record-input-source", targetId = "record-input-target";
        ObjectNode source = (ObjectNode) snapshot(sourceId);
        ObjectNode cells = (ObjectNode) source.path("sheets").get(0).path("cells"); cells.removeAll();
        cells.putObject("0").putObject("0").put("value", "ID");
        ObjectNode row = cells.putObject("1"); row.putObject("0").put("value", "r1"); row.putObject("1").put("value", 2);
        ((com.fasterxml.jackson.databind.node.ArrayNode) source.path("dataModel").path("tables")).add(mapper.readTree("""
          {"id":"records","name":"Records","sourceSheetId":"sheet-1","recordIdFieldId":"id","rowCount":1,"blockSize":1024,"blocks":[],"revision":0,
           "sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":2},
           "fields":[{"id":"id","name":"ID","ordinal":0,"type":"text"},{"id":"qty","name":"Qty","ordinal":1,"type":"number"},{"id":"total","name":"Total","ordinal":2,"type":"number","calculation":{"kind":"formula","formula":"=[@qty]*2"}}]}
          """));
        catalog.create(new CreateWorkbookRequest(sourceId, "Range access", source), owner);
        ObjectNode target = (ObjectNode) snapshot(targetId);
        ((com.fasterxml.jackson.databind.node.ArrayNode) target.path("dataModel").path("externalLinks")).add(mapper.readTree("""
          {"id":"record-link","token":"Records.xlsx","sourceUnitId":"record-input-source","sheets":[{"token":"Sales","sheetId":"sheet-1"}]}
          """));
        catalog.create(new CreateWorkbookRequest(targetId, "Range access", target), owner);
        acl.grant(sourceId, owner, reader, com.xc.luckysheet.server.contract.WorkbookRole.VIEWER);
        acl.grant(targetId, owner, reader, com.xc.luckysheet.server.contract.WorkbookRole.VIEWER);
        assertEquals(0, node(operations.readExternalCalculationGraph(targetId, reader, List.of()), sourceId).path("revision").asLong());
        rangeAccess.create(sourceId, owner, new RangeAccessRegionRequest("sheet-1", new RangeRef("sheet-1", 1, 1, 1, 1), RangeAccessLevel.HIDDEN, List.of()));
        JsonNode denied = node(operations.readExternalCalculationGraph(targetId, reader, List.of()), sourceId);
        assertEquals("denied", denied.path("state").asText()); assertFalse(denied.has("snapshot"));
        assertTrue(denied.path("error").path("message").asText().contains("unreadable inputs"));
        assertEquals(0, node(operations.readExternalCalculationGraph(targetId, owner, List.of()), sourceId).path("revision").asLong());
    }

    @Test
    void multiWorkbookClosurePreservesVersionsAndRejectsCircularBindingsWithoutWrites() throws Exception {
        String owner = "graph-owner", a = "graph-a", b = "graph-b", c = "graph-c";
        catalog.create(new CreateWorkbookRequest(a, "Range access", snapshot(a)), owner);
        ObjectNode middle = (ObjectNode) snapshot(b);
        ((com.fasterxml.jackson.databind.node.ArrayNode) middle.path("dataModel").path("externalLinks")).add(binding(a, "A.xlsx"));
        catalog.create(new CreateWorkbookRequest(b, "Range access", middle), owner);
        ObjectNode target = (ObjectNode) snapshot(c);
        ((com.fasterxml.jackson.databind.node.ArrayNode) target.path("dataModel").path("externalLinks")).add(binding(b, "B.xlsx"));
        catalog.create(new CreateWorkbookRequest(c, "Range access", target), owner);
        JsonNode graph = operations.readExternalCalculationGraph(c, owner, List.of());
        assertEquals(3, graph.path("nodes").size());
        for (String id : List.of(a, b, c)) { assertEquals("connected", node(graph, id).path("state").asText()); assertEquals(0, node(graph, id).path("revision").asLong()); }
        JsonNode preimage = operations.readSnapshot(a, owner).snapshot();
        ServiceException preflight = assertThrows(ServiceException.class, () -> operations.validateExternalBinding(a, binding(c, "C.xlsx"), owner, List.of()));
        assertEquals("CIRCULAR_DEPENDENCY", preflight.code());
        ObjectNode params = mapper.createObjectNode(); params.set("link", binding(c, "C.xlsx"));
        var operation = new OperationEnvelope("graph-client", OperationEnvelope.SCHEMA, "graph-cycle-rejected", a, 1, 0,
                List.of(new OperationMutation("externalLink.set", "sheet-1", params)), Instant.now());
        ServiceException forged = assertThrows(ServiceException.class, () -> operations.commit(a, operation, owner));
        assertEquals("CIRCULAR_DEPENDENCY", forged.code());
        assertEquals(preimage, operations.readSnapshot(a, owner).snapshot());
        for (String id : List.of(a, b, c)) assertEquals(0, operations.readSnapshot(id, owner).revision());
        assertThrows(ServiceException.class, () -> operations.operationResult(a, operation.operationId(), owner));
    }

    private JsonNode binding(String source, String token) throws Exception {
        return mapper.readTree("{\"id\":\"" + token + "\",\"token\":\"" + token + "\",\"sourceUnitId\":\"" + source
                + "\",\"sheets\":[{\"token\":\"Sheet1\",\"sheetId\":\"sheet-1\"}]}");
    }

    @Test
    void concurrentOppositeBindingsCommitOnlyOneDirectionAndKeepRejectedHistoryEmpty() throws Exception {
        String owner = "concurrent-graph-owner", a = "concurrent-graph-a", b = "concurrent-graph-b";
        catalog.create(new CreateWorkbookRequest(a, "Range access", snapshot(a)), owner);
        catalog.create(new CreateWorkbookRequest(b, "Range access", snapshot(b)), owner);
        var start = new java.util.concurrent.CountDownLatch(1);
        try (var executor = java.util.concurrent.Executors.newFixedThreadPool(2)) {
            var results = new java.util.ArrayList<java.util.concurrent.Future<String>>();
            for (String target : List.of(a, b)) {
                ObjectNode params = mapper.createObjectNode(); params.set("link", binding(target.equals(a) ? b : a, target + ".xlsx"));
                var operation = new OperationEnvelope("concurrent-client-" + target, OperationEnvelope.SCHEMA, "concurrent-bind-" + target,
                        target, 1, 0, List.of(new OperationMutation("externalLink.set", "sheet-1", params)), Instant.now());
                results.add(executor.submit(() -> {
                    if (!start.await(10, java.util.concurrent.TimeUnit.SECONDS)) throw new AssertionError("Concurrent start did not arrive");
                    try { operations.commit(target, operation, owner); return "COMMITTED"; }
                    catch (ServiceException error) { return error.code(); }
                }));
            }
            start.countDown();
            String first = results.get(0).get(30, java.util.concurrent.TimeUnit.SECONDS), second = results.get(1).get(30, java.util.concurrent.TimeUnit.SECONDS);
            assertEquals(1, java.util.stream.Stream.of(first, second).filter("COMMITTED"::equals).count());
            assertEquals(1, java.util.stream.Stream.of(first, second).filter("CIRCULAR_DEPENDENCY"::equals).count());
            String accepted = first.equals("COMMITTED") ? a : b, rejected = accepted.equals(a) ? b : a;
            assertEquals(1, operations.readSnapshot(accepted, owner).revision());
            assertEquals(0, operations.readSnapshot(rejected, owner).revision());
            assertTrue(operations.readSnapshot(rejected, owner).snapshot().path("dataModel").path("externalLinks").isEmpty());
            assertThrows(ServiceException.class, () -> operations.operationResult(rejected, "concurrent-bind-" + rejected, owner));
            assertEquals(2, operations.readExternalCalculationGraph(accepted, owner, List.of()).path("nodes").size());
        }
    }

    private JsonNode node(JsonNode graph, String id) {
        for (JsonNode node : graph.path("nodes")) if (node.path("unitId").asText().equals(id)) return node;
        throw new AssertionError("Missing calculation node " + id);
    }

    @Test
    void graphDiscardsAnUnlockedCachedEntityBeforeCapturingTheCommittedSourceVersion() throws Exception {
        String owner = "cached-graph-owner", source = "cached-graph-source", target = "cached-graph-target";
        catalog.create(new CreateWorkbookRequest(source, "Range access", snapshot(source)), owner);
        ObjectNode root = (ObjectNode) snapshot(target);
        ((com.fasterxml.jackson.databind.node.ArrayNode) root.path("dataModel").path("externalLinks")).add(binding(source, "Source.xlsx"));
        catalog.create(new CreateWorkbookRequest(target, "Range access", root), owner);
        ObjectNode params = mapper.createObjectNode().put("row", 0).put("column", 0); params.putObject("value").put("value", 40);
        ObjectNode authority = params.putObject("writeAuthority").put("kind", "script");
        authority.putObject("target").put("sheetId", "sheet-1").put("row", 0).put("column", 0);
        authority.set("candidate", params.path("value").deepCopy()); authority.putObject("validationDecision").put("status", "accepted");
        var write = new OperationEnvelope("cached-client", OperationEnvelope.SCHEMA, "cached-source-write", source, 1, 0,
                List.of(new OperationMutation("cell.set", "sheet-1", params)), Instant.now());
        try (var writer = java.util.concurrent.Executors.newSingleThreadExecutor()) {
            new org.springframework.transaction.support.TransactionTemplate(transactionManager).executeWithoutResult(transaction -> {
                assertEquals(0, store.find(source).orElseThrow().revision());
                try { writer.submit(() -> operations.commit(source, write, owner)).get(15, java.util.concurrent.TimeUnit.SECONDS); }
                catch (Exception error) { throw new AssertionError("Independent source write failed", error); }
                JsonNode committed = node(operations.readExternalCalculationGraph(target, owner, List.of()), source);
                assertEquals(1, committed.path("revision").asLong());
                assertEquals(40, committed.path("snapshot").path("sheets").get(0).path("cells").path("0").path("0").path("value").asInt());
            });
        }
    }

    @Test
    void historyRestoreRevalidatesCurrentTopologyAndLifecycleSourcesNeverExposeSnapshots() throws Exception {
        String owner = "restore-graph-owner", a = "restore-graph-a", b = "restore-graph-b";
        catalog.create(new CreateWorkbookRequest(a, "Range access", snapshot(a)), owner);
        catalog.create(new CreateWorkbookRequest(b, "Range access", snapshot(b)), owner);
        ObjectNode set = mapper.createObjectNode(); set.set("link", binding(b, "B.xlsx"));
        operations.commit(a, new OperationEnvelope("restore-client", OperationEnvelope.SCHEMA, "restore-a-bind", a, 1, 0,
                List.of(new OperationMutation("externalLink.set", "sheet-1", set)), Instant.now()), owner);
        operations.commit(a, new OperationEnvelope("restore-client", OperationEnvelope.SCHEMA, "restore-a-unbind", a, 2, 1,
                List.of(new OperationMutation("externalLink.remove", "sheet-1", mapper.createObjectNode().put("linkId", "B.xlsx"))), Instant.now()), owner);
        ObjectNode opposite = mapper.createObjectNode(); opposite.set("link", binding(a, "A.xlsx"));
        operations.commit(b, new OperationEnvelope("restore-client", OperationEnvelope.SCHEMA, "restore-b-bind", b, 1, 0,
                List.of(new OperationMutation("externalLink.set", "sheet-1", opposite)), Instant.now()), owner);
        JsonNode before = operations.readSnapshot(a, owner).snapshot();
        ServiceException error = assertThrows(ServiceException.class, () -> operations.restore(a, new com.xc.luckysheet.server.contract.RestoreRequest(1, "Circular historical binding"), owner));
        assertEquals("CIRCULAR_DEPENDENCY", error.code()); assertEquals(2, operations.readSnapshot(a, owner).revision());
        assertEquals(before, operations.readSnapshot(a, owner).snapshot());
        catalog.moveToTrash(a, owner);
        JsonNode trashed = node(operations.readExternalCalculationGraph(b, owner, List.of()), a);
        assertEquals("broken", trashed.path("state").asText()); assertFalse(trashed.has("snapshot"));
        catalog.restoreFromTrash(a, owner);
        assertEquals("connected", node(operations.readExternalCalculationGraph(b, owner, List.of()), a).path("state").asText());
        catalog.moveToTrash(a, owner); catalog.purge(a, owner);
        JsonNode purged = node(operations.readExternalCalculationGraph(b, owner, List.of()), a);
        assertEquals("broken", purged.path("state").asText()); assertFalse(purged.has("snapshot"));
    }

    private JsonNode snapshot(String unitId) throws Exception {
        ObjectNode root = (ObjectNode) mapper.readTree("""
                {"schema":"WorkbookSnapshot","version":11,"unitId":"%s","name":"Range access",
                 "dimensionMetrics":{"normalFontFamily":"Calibri","normalFontSizePx":14.6666666667,"maximumDigitWidthPx":7},
                 "calculationSettings":{},"editingOptions":{"allowEditDirectly":true,"moveAfterEnter":true,"enterDirection":"down","formulaAutoComplete":true,"valueAutoComplete":true,"fixedDecimalPlaces":null},
                 "dataModel":{"externalLinks":[],"sources":[],"tables":[],"relationships":[],"views":[]},
                 "sheets":[{"kind":"worksheet","id":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,
                   "cells":{"0":{"0":{"value":"VISIBLE-CELL-VALUE"},"1":{"value":"SECRET-CELL-VALUE"}}},
                   "merges":[],"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                   "pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"hyperlinks":[],
                   "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}]
                }
                """.formatted(unitId));
        return root;
    }
}
