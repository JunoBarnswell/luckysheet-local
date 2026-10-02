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
        "DATABASE_URL=jdbc:h2:mem:range_access_authorization;DB_CLOSE_DELAY=-1",
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
class RangeAccessAuthorizationIntegrationTest {
    @Autowired private WorkbookCatalogService catalog;
    @Autowired private GuestShareService shares;
    @Autowired private RangeAccessService rangeAccess;
    @Autowired private WorkbookOperationService operations;
    @Autowired private ObjectMapper mapper;

    @Test
    void hiddenSnapshotDataIsWithheldAndHiddenWritesAreRejectedWhileVisibleWritesSucceed() throws Exception {
        String unitId = "range-access-boundary";
        String owner = "range-owner";
        catalog.create(new CreateWorkbookRequest(unitId, "Range access", snapshot(unitId)), owner);
        var share = shares.create(unitId, new ShareCreateRequest("editor", Instant.now().plusSeconds(600)), owner);
        String guest = "guest:" + share.shareId();

        rangeAccess.create(unitId, owner, new RangeAccessRegionRequest("sheet-1",
                new RangeRef("sheet-1", 0, 0, 1, 1), RangeAccessLevel.HIDDEN, List.of()));

        JsonNode safeSnapshot = operations.readSnapshot(unitId, guest).snapshot();
        String responseJson = mapper.writeValueAsString(safeSnapshot);
        assertFalse(responseJson.contains("SECRET-CELL-VALUE"));
        assertFalse(safeSnapshot.path("sheets").get(0).path("cells").path("0").has("1"));

        ServiceException denied = assertThrows(ServiceException.class,
                () -> operations.commit(unitId, operation(unitId, "hidden-write", 1, 0, 1, "ATTACK"), guest));
        assertEquals("ACCESS_DENIED", denied.code());
        assertEquals(0, operations.readSnapshot(unitId, owner).revision());
        assertEquals("SECRET-CELL-VALUE", operations.readSnapshot(unitId, owner).snapshot()
                .path("sheets").get(0).path("cells").path("0").path("1").path("value").asText());

        operations.commit(unitId, operation(unitId, "visible-write", 2, 0, 0, "VISIBLE-UPDATED"), guest);
        assertEquals("VISIBLE-UPDATED", operations.readSnapshot(unitId, owner).snapshot()
                .path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());

        OperationEnvelope formulaWrite = new OperationEnvelope("range-access-session", OperationEnvelope.SCHEMA,
                "visible-formula-hidden-source", unitId, 3, 1,
                List.of(new OperationMutation("cell.set", "sheet-1", mapper.readTree(
                        "{\"row\":0,\"column\":0,\"value\":{\"formula\":\"=B1\",\"value\":\"SECRET-CELL-VALUE\"}}"))), Instant.now());
        ServiceException formulaDenied = assertThrows(ServiceException.class,
                () -> operations.commit(unitId, formulaWrite, guest));
        assertEquals("ACCESS_HIDDEN", formulaDenied.code());
        assertEquals(1, operations.readSnapshot(unitId, owner).revision());
    }

    @Test
    void overlappingRegionsAreRejectedAndDoNotAdvanceAccessRevision() throws Exception {
        String unitId = "range-access-overlap";
        String owner = "range-overlap-owner";
        catalog.create(new CreateWorkbookRequest(unitId, "Range access", snapshot(unitId)), owner);
        rangeAccess.create(unitId, owner, new RangeAccessRegionRequest("sheet-1",
                new RangeRef("sheet-1", 0, 2, 0, 2), RangeAccessLevel.READ, List.of()));
        long revision = rangeAccess.projection(unitId, owner, List.of()).accessRevision();

        ServiceException overlap = assertThrows(ServiceException.class, () -> rangeAccess.create(unitId, owner,
                new RangeAccessRegionRequest("sheet-1", new RangeRef("sheet-1", 2, 3, 2, 3),
                        RangeAccessLevel.HIDDEN, List.of())));

        assertEquals("ACCESS_REGION_OVERLAP", overlap.code());
        assertEquals(revision, rangeAccess.projection(unitId, owner, List.of()).accessRevision());
        assertEquals(1, rangeAccess.list(unitId, owner).size());
    }

    @Test
    void structuralChangesReuseCanonicalRangeTransformAndFollowSheetLifecycle() throws Exception {
        String unitId = "range-access-structure";
        String owner = "range-structure-owner";
        catalog.create(new CreateWorkbookRequest(unitId, "Range access", snapshot(unitId)), owner);
        rangeAccess.create(unitId, owner, new RangeAccessRegionRequest("sheet-1",
                new RangeRef("sheet-1", 1, 3, 2, 5), RangeAccessLevel.EDIT, List.of()));

        rangeAccess.applyStructuralMutation(unitId,
                new OperationMutation("rows.inserted", "sheet-1", mapper.readTree("{\"at\":1,\"count\":1}")),
                snapshot(unitId));
        assertEquals(new RangeRef("sheet-1", 2, 4, 2, 5), rangeAccess.list(unitId, owner).getFirst().range());

        rangeAccess.applyStructuralMutation(unitId,
                new OperationMutation("rows.deleted", "sheet-1", mapper.readTree("{\"at\":0,\"count\":1}")),
                snapshot(unitId));
        assertEquals(new RangeRef("sheet-1", 1, 3, 2, 5), rangeAccess.list(unitId, owner).getFirst().range());

        ObjectNode copiedSheet = ((ObjectNode) snapshot(unitId).path("sheets").get(0)).deepCopy();
        copiedSheet.put("id", "sheet-copy").put("name", "Sheet1 Copy");
        ObjectNode duplicatedSnapshot = (ObjectNode) snapshot(unitId);
        ((com.fasterxml.jackson.databind.node.ArrayNode) duplicatedSnapshot.path("sheets")).add(copiedSheet);
        rangeAccess.applyStructuralMutation(unitId,
                new OperationMutation("sheet.duplicated", "sheet-1",
                        mapper.readTree("{\"sourceSheetId\":\"sheet-1\",\"newId\":\"sheet-copy\"}")),
                duplicatedSnapshot);
        assertEquals(2, rangeAccess.list(unitId, owner).size());
        assertTrue(rangeAccess.list(unitId, owner).stream().anyMatch(region ->
                "sheet-copy".equals(region.sheetId())
                        && new RangeRef("sheet-copy", 1, 3, 2, 5).equals(region.range())));

        rangeAccess.applyStructuralMutation(unitId,
                new OperationMutation("sheet.remove", "sheet-copy", mapper.readTree("{\"id\":\"sheet-copy\"}")),
                duplicatedSnapshot);
        assertEquals(1, rangeAccess.list(unitId, owner).size());
        assertEquals("sheet-1", rangeAccess.list(unitId, owner).getFirst().sheetId());
    }

    private OperationEnvelope operation(String unitId, String operationId, long sequence,
                                        int row, int column, String value) {
        ObjectNode cell = mapper.createObjectNode().put("value", value);
        ObjectNode target = mapper.createObjectNode().put("sheetId", "sheet-1").put("row", row).put("column", column);
        ObjectNode decision = mapper.createObjectNode().put("status", "accepted");
        ObjectNode authority = mapper.createObjectNode().put("kind", "direct-entry");
        authority.set("target", target.deepCopy());
        authority.set("candidate", cell.deepCopy());
        authority.set("validationDecision", decision);
        ObjectNode params = mapper.createObjectNode().put("row", row).put("column", column);
        params.set("value", cell);
        params.set("writeAuthority", authority);
        return new OperationEnvelope("range-access-session", OperationEnvelope.SCHEMA, operationId, unitId,
                sequence, 0, List.of(new OperationMutation("cell.set", "sheet-1", params)), Instant.now());
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
