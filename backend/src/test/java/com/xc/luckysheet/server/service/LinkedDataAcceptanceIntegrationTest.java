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
        assertThrows(ServiceException.class, () -> operations.readExternalLink(targetId, "link-1", guest, List.of()));
        JsonNode input = operations.readExternalLink(targetId, "link-1", owner, List.of());
        com.xc.luckysheet.server.contract.WorkbookSnapshotValidator.requireCanonical(input.path("snapshot"), sourceId);
        assertEquals(owner, input.path("subject").asText());
        assertEquals(0, input.path("sourceRevision").asLong());
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
        var result = operations.commit(unit, operation, owner);
        JsonNode persisted = operations.readSnapshot(unit, owner).snapshot();
        assertEquals("=SUM(Sheet2:Sheet3!A1)", persisted.path("sheets").get(2).path("cells").path("0").path("0").path("formula").asText());
        assertEquals(1, operations.readSnapshot(unit, owner).revision());
        assertEquals(10, result.operation().mutations().getFirst().structuralPatch().version());
        assertEquals(1, result.operation().mutations().getFirst().structuralPatch().formulaOwnerDeltas().size());
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
