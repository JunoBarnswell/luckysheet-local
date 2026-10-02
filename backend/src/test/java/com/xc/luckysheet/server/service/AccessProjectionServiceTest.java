package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationOrigin;
import com.xc.luckysheet.server.contract.RangeAccessLevel;
import com.xc.luckysheet.server.contract.RangeAccessRegion;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class AccessProjectionServiceTest {
    private final ObjectMapper mapper = new ObjectMapper().findAndRegisterModules();
    private final AccessProjectionService projection = new AccessProjectionService(mapper);

    @Test
    void revisionPayloadIsWithheldWhenVisibleFormulaDependsOnHiddenCell() throws Exception {
        RangeAccessRegion hidden = new RangeAccessRegion("region-1", "book-1", "sheet-1",
                new RangeRef("sheet-1", 0, 0, 1, 1), RangeAccessLevel.HIDDEN, List.of(),
                "owner", Instant.EPOCH, Instant.EPOCH);
        var access = resolver(hidden);
        JsonNode params = mapper.readTree("{\"row\":0,\"column\":2,\"value\":{\"formula\":\"=B1\",\"value\":1234}}");
        var mutation = new CommittedOperationMutation("cell.set", "sheet-1", params,
                List.of(new RangeRef("sheet-1", 0, 0, 2, 2)));
        var operation = new CommittedOperationEnvelope("session", OperationEnvelope.SCHEMA, "op-1", "book-1",
                "editor", OperationOrigin.CLIENT, 1, 0, 1, List.of(mutation), Instant.EPOCH, Instant.EPOCH);

        assertFalse(projection.canDeliver(operation, access));
    }

    @Test
    void snapshotProjectionRemovesHiddenCellsAndRedactsFormulaAndDrawingReferences() throws Exception {
        RangeAccessRegion hidden = new RangeAccessRegion("region-1", "book-1", "sheet-1",
                new RangeRef("sheet-1", 0, 0, 1, 1), RangeAccessLevel.HIDDEN, List.of(),
                "owner", Instant.EPOCH, Instant.EPOCH);
        RangeAccessResolver access = resolver(hidden);
        JsonNode source = mapper.readTree("""
                {"schema":"WorkbookSnapshot","version":10,"unitId":"book-1",
                 "dataModel":{"sources":[{"id":"hidden-source","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":1,"endColumn":1},"fields":[{"id":"f1","name":"Secret Column"}],"blocks":[{"id":"hidden-block","checksum":"secret-checksum"}]},{"id":"region-only-hidden-source","fields":[{"id":"f3","name":"Region Secret"}],"blocks":[{"id":"region-hidden-block","checksum":"region-secret-checksum"}]},{"id":"visible-source","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":2,"endColumn":2},"fields":[{"id":"f2","name":"Visible Column"}],"blocks":[]}],
                   "tables":[{"id":"hidden-table","sourceId":"hidden-source","fields":[{"id":"f1","name":"Secret Column"}]},{"id":"region-hidden-table","sourceId":"region-only-hidden-source","fields":[{"id":"f3","name":"Region Secret"}]},{"id":"visible-table","sourceId":"visible-source","fields":[{"id":"f2","name":"Visible Column"}]}],
                   "relationships":[{"id":"hidden-link","fromTableId":"hidden-table","toTableId":"public-table"}],
                   "views":[{"id":"hidden-view","tableId":"hidden-table","filters":[{"values":["secret-filter"]}]},{"id":"visible-view","tableId":"visible-table"}]},
                 "queryDefinitions":[{"id":"hidden-query","lastTarget":{"kind":"range","sheetId":"sheet-1","range":{"startRow":0,"endRow":0,"startColumn":1,"endColumn":1}}},{"id":"region-hidden-query","lastTarget":{"kind":"workbook-table","tableId":"region-hidden-table"}},{"id":"visible-query","lastTarget":{"kind":"range","sheetId":"sheet-1","range":{"startRow":0,"endRow":0,"startColumn":2,"endColumn":2}}}],
                 "definedNameModels":[{"name":"SensitiveRange","formula":"=Sheet1!B1","scope":"workbook"}],"definedNames":{"SensitiveRange":"=Sheet1!B1"},
                 "sheets":[{
                  "id":"sheet-1","name":"Sheet1","cells":{"0":{
                    "0":{"value":"visible"},
                    "1":{"value":"secret"},
                    "2":{"formula":"=B1","value":42,"displayValue":"42","richText":[{"text":"cached-secret"}],"formulaValue":{"kind":"number","value":42},"formulaMetadata":{"kind":"normal","sourceFormula":"=B1"},"presentation":{"kind":"barcode","source":{"kind":"formula","formula":"=B1"}}},
                    "3":{"formula":"=SUM(B1)","value":84,"formulaValue":{"kind":"number","value":84}},
                    "4":{"formula":"=\\\"B1\\\"","value":"B1","formulaValue":{"kind":"text","value":"B1"}}
                  }},
                  "drawings":[{"id":"chart-1","sheetId":"sheet-1","payloadId":"payload-1","anchor":{"row":0,"column":0}},{"id":"pivot-chart","sheetId":"sheet-1","payloadId":"pivot-payload","anchor":{"row":0,"column":4}},{"id":"visible-pivot-chart","sheetId":"sheet-1","payloadId":"visible-pivot-payload","anchor":{"row":0,"column":5}}],
                  "drawingPayloads":{"payload-1":{"kind":"chart","series":[{"range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":1,"endColumn":1}}]},"pivot-payload":{"kind":"chart","source":{"kind":"pivot","pivotId":"pivot-source"}},"visible-pivot-payload":{"kind":"chart","source":{"kind":"pivot","pivotId":"pivot-visible"}},"orphan":{"kind":"shape"}},
                  "dataRegions":[{"id":"hidden-region","sourceId":"region-only-hidden-source","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":1,"endColumn":1},"headerRow":0}],
                  "sheetTables":[{"id":"hidden-sheet-table","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":1,"endColumn":1}},{"id":"visible-sheet-table","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":2,"endColumn":2}}],
                  "pivots":[
                    {"id":"pivot-source","source":{"kind":"data-source","dataSourceId":"hidden-source"},"fieldCatalog":{"fields":[{"fieldId":"f1","name":"Secret Column","values":["SECRET-PIVOT-MEMBER"]}]}},
                    {"id":"pivot-region-source","source":{"kind":"data-source","dataSourceId":"region-only-hidden-source"},"fieldCatalog":{"fields":[{"fieldId":"f3","name":"Region Secret","values":["REGION-PIVOT-SECRET"]}]}},
                    {"id":"pivot-table","source":{"kind":"table","tableId":"hidden-sheet-table"},"fieldCatalog":{"fields":[{"fieldId":"f1","name":"Private Table","values":["SECRET-TABLE-MEMBER"]}]}},
                    {"id":"pivot-name","source":{"kind":"named-range","name":"SensitiveRange"},"fieldCatalog":{"fields":[{"fieldId":"f1","name":"Named Source","values":["SECRET-NAME-MEMBER"]}]}},
                    {"id":"pivot-visible","source":{"kind":"data-source","dataSourceId":"visible-source"},"fieldCatalog":{"fields":[{"fieldId":"f2","name":"Visible Column","values":["public-member"]}]}}],
                  "conditionalFormats":[
                    {"id":"cf-hidden-target","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":1,"endColumn":1}],"type":"highlight","value1":"secret-threshold"},
                    {"id":"cf-hidden-formula","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":3,"endColumn":3}],"type":"highlight","operator":"formula","value1":"=B1"},
                    {"id":"cf-visible","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":1,"endRow":1,"startColumn":3,"endColumn":3}],"type":"highlight","value1":"public-threshold"}],
                  "dataValidations":[
                    {"id":"dv-hidden-target","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":1,"endColumn":1}],"type":"list","listSource":{"kind":"values","values":["secret-option"]}},
                    {"id":"dv-hidden-formula","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":1,"endRow":1,"startColumn":3,"endColumn":3}],"type":"list","listSource":{"kind":"formula","formula":"=B1"}}],
                  "review":{"notesByCell":{"0:1":{"text":"secret note"},"0:0":{"text":"visible note"}},"threadIdsByCell":{"0:1":["hidden-thread"]},"threadsById":{"hidden-thread":{"body":"secret discussion"}}},
                  "hyperlinks":[{"row":0,"column":1,"target":"https://secret.invalid"},{"row":0,"column":0,"target":"https://public.invalid"}]
                }]}
                """);

        JsonNode projected = projection.projectSnapshot(source, access);
        JsonNode sheet = projected.path("sheets").get(0);
        JsonNode cells = sheet.path("cells").path("0");

        assertEquals("visible", cells.path("0").path("value").asText());
        assertFalse(cells.has("1"));
        assertEquals("=#BLOCKED!", cells.path("2").path("formula").asText());
        assertTrue(cells.path("2").get("value").isNull());
        assertFalse(cells.path("2").has("displayValue"));
        assertFalse(cells.path("2").has("richText"));
        assertFalse(cells.path("2").has("formulaMetadata"));
        assertFalse(cells.path("2").has("presentation"));
        assertEquals("#BLOCKED!", cells.path("2").path("formulaValue").path("code").asText());
        assertEquals("=#BLOCKED!", cells.path("3").path("formula").asText());
        assertEquals("=\"B1\"", cells.path("4").path("formula").asText());
        assertEquals(1, sheet.path("drawings").size());
        assertEquals("visible-pivot-chart", sheet.path("drawings").get(0).path("id").asText());
        assertEquals(1, sheet.path("drawingPayloads").size());
        assertEquals(1, sheet.path("pivots").size());
        assertEquals("pivot-visible", sheet.path("pivots").get(0).path("id").asText());
        assertEquals(1, sheet.path("sheetTables").size());
        assertEquals("visible-sheet-table", sheet.path("sheetTables").get(0).path("id").asText());
        assertEquals(1, projected.path("dataModel").path("sources").size());
        assertEquals("visible-source", projected.path("dataModel").path("sources").get(0).path("id").asText());
        assertEquals(1, projected.path("dataModel").path("tables").size());
        assertEquals(0, projected.path("dataModel").path("relationships").size());
        assertEquals(1, projected.path("dataModel").path("views").size());
        assertEquals("visible-view", projected.path("dataModel").path("views").get(0).path("id").asText());
        assertEquals(1, projected.path("queryDefinitions").size());
        assertEquals("visible-query", projected.path("queryDefinitions").get(0).path("id").asText());
        assertEquals(0, projected.path("definedNameModels").size());
        assertFalse(projected.path("definedNames").has("SensitiveRange"));
        assertFalse(sheet.path("review").path("notesByCell").has("0:1"));
        assertFalse(sheet.path("review").path("threadsById").has("hidden-thread"));
        assertEquals(1, sheet.path("conditionalFormats").size());
        assertEquals("cf-visible", sheet.path("conditionalFormats").get(0).path("id").asText());
        assertEquals(0, sheet.path("dataValidations").size());
        assertEquals(1, sheet.path("hyperlinks").size());
        assertEquals("https://public.invalid", sheet.path("hyperlinks").get(0).path("target").asText());
        assertEquals("secret", source.path("sheets").get(0).path("cells").path("0").path("1").path("value").asText(),
                "projection must not mutate the canonical snapshot");
    }

    private RangeAccessResolver resolver(RangeAccessRegion region) {
        List<RangeAccessRegion> regions = List.of(region);
        return new RangeAccessResolver(new RangeAccessIndex(regions), regions,
                new RangeAccessContext("guest", WorkbookAclRole.EDITOR, List.of(), 3));
    }
}
