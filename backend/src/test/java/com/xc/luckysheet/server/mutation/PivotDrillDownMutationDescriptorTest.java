package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class PivotDrillDownMutationDescriptorTest {
    private final ObjectMapper mapper = new ObjectMapper();
    private final PivotDrillDownMutationDescriptor descriptor = new PivotDrillDownMutationDescriptor("pivot.drilldown.add");

    @Test
    void groupsInterleavedProvenanceByRecordAndSourceIncludingUnmatchedLeftRows() throws Exception {
        ObjectNode snapshot = snapshot();
        ObjectNode params = params("""
                [{"sourceId":"orders","recordId":"a","sheetId":"sheet-1","row":1},
                 {"sourceId":"orders","recordId":"b","sheetId":"sheet-1","row":2},
                 {"sourceId":"customers","recordId":"a","sheetId":"sheet-1","row":2}]
                """);
        JsonNode result = descriptor.apply(snapshot, mutation(params));
        JsonNode cells = result.path("sheets").get(1).path("cells");
        assertEquals("CustomerId", cells.path("0").path("0").path("value").asText());
        assertTrue(cells.path("1").isMissingNode());
        assertEquals(2, result.path("dataModel").path("sources").get(0).path("rowCount").asInt());
        assertEquals("detail-source", result.path("sheets").get(1).path("dataRegions").get(0).path("sourceId").asText());
        assertEquals(1, snapshot.path("sheets").size());
    }

    @Test
    void rejectsIncompleteInnerJoinAndMissingLeftRoot() throws Exception {
        ObjectNode snapshot = snapshot();
        ObjectNode params = params("[{\"sourceId\":\"customers\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":2}]");
        assertThrows(ServiceException.class, () -> descriptor.apply(snapshot, mutation(params)));
        ((ObjectNode) snapshot.path("sheets").get(0).path("pivots").get(0).path("source").path("relationships").get(0)).put("join", "inner");
        ObjectNode incomplete = params("[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1}]");
        assertThrows(ServiceException.class, () -> descriptor.apply(snapshot, mutation(incomplete)));
        assertEquals(1, snapshot.path("sheets").size());
    }

    @Test
    void rejectsOutsideRangeUnknownDuplicateAndIdentitylessPathsBeforeCreatingTarget() throws Exception {
        ObjectNode snapshot = snapshot();
        for (String paths : new String[] {
                "[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":0}]",
                "[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":3}]",
                "[{\"sourceId\":\"unknown\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1}]",
                "[{\"sheetId\":\"sheet-1\",\"row\":1}]",
                "[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1},{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":2}]"
        }) {
            OperationMutation mutation = mutation(params(paths));
            assertThrows(ServiceException.class, () -> descriptor.affectedRanges(snapshot, mutation));
            assertThrows(ServiceException.class, () -> descriptor.apply(snapshot, mutation));
        }
        assertEquals(1, snapshot.path("sheets").size());
    }

    @Test
    void rejectsAnchorIntegerOverflow() throws Exception {
        ObjectNode snapshot = snapshot();
        ObjectNode params = params("[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1}]");
        ((ObjectNode) params.get("target")).put("row", 4_294_967_296L);
        assertThrows(ServiceException.class, () -> descriptor.apply(snapshot, mutation(params)));
        assertEquals(1, snapshot.path("sheets").size());
    }

    @Test
    void blockBackedWorksheetRangesPublishThePreparedDetailSourceInsteadOfBlankCells() throws Exception {
        ObjectNode snapshot = snapshot();
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        JsonNode sourceRange = sheet.path("pivots").get(0).path("source").path("ranges").get(0).path("range").deepCopy();
        ObjectNode manifest = snapshot.putObject("dataModel").putArray("sources").addObject();
        manifest.put("schema", "DataSourceManifest").put("version", 1).put("id", "blocks").put("name", "Blocks")
                .put("kind", "chunked-table").put("sourceSheetId", "sheet-1").put("rowCount", 2).put("blockRowCount", 65536).put("revision", 0);
        manifest.set("sourceRange", sourceRange);
        var fields = manifest.putArray("fields");
        fields.addObject().put("id", "customer").put("name", "CustomerId").put("ordinal", 0).put("type", "text");
        fields.addObject().put("id", "amount").put("name", "Amount").put("ordinal", 1).put("type", "number");
        manifest.putArray("blocks").addObject().put("id", "block").put("dataSourceId", "blocks")
                .put("startRow", 0).put("rowCount", 2).put("storageKey", "block").put("checksum", "a".repeat(64))
                .put("byteLength", 1).put("encoding", "columnar-v1").put("revision", 0);
        for (String row : new String[] {"1", "2"}) {
            ((ObjectNode) sheet.path("cells").path(row)).remove("0");
            ((ObjectNode) sheet.path("cells").path(row)).remove("1");
        }
        ObjectNode region = sheet.putArray("dataRegions").addObject();
        region.put("id", "region").put("sourceId", "blocks").put("headerRow", 0).put("revision", 0);
        region.set("range", sourceRange.deepCopy());
        ObjectNode params = params("[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1}]");
        OperationMutation mutation = mutation(params);
        JsonNode result = descriptor.apply(snapshot, mutation);
        assertEquals(2, result.path("sheets").size());
        assertEquals("detail-source", result.path("sheets").get(1).path("dataRegions").get(0).path("sourceId").asText());
        assertTrue(result.path("sheets").get(1).path("cells").path("1").isMissingNode());
    }

    @Test
    void usesEvaluatedHeadersWithoutReadingDetailValuesFromSparseCells() throws Exception {
        ObjectNode snapshot = snapshot();
        ObjectNode cell = (ObjectNode) snapshot.path("sheets").get(0).path("cells").path("1").path("1");
        ((ObjectNode) snapshot.path("sheets").get(0).path("cells").path("0").path("1")).put("formulaValue", "Computed Amount");
        cell.putObject("formulaValue").put("kind", "error").put("code", "#DIV/0!");
        ObjectNode params = params("[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1}]");
        ((ObjectNode) params.path("detail").path("source").path("fields").get(1)).put("name", "Computed Amount");
        ((com.fasterxml.jackson.databind.node.ArrayNode) params.path("detail").path("headers")).set(1, mapper.getNodeFactory().textNode("Computed Amount"));
        OperationMutation mutation = mutation(params);
        JsonNode result = descriptor.apply(snapshot, mutation);
        assertEquals("Computed Amount", result.path("sheets").get(1).path("cells").path("0").path("1").path("value").asText());
        assertTrue(result.path("sheets").get(1).path("cells").path("1").isMissingNode());
        cell.putObject("formulaValue").put("unexpected", true);
        JsonNode repeated = descriptor.apply(snapshot, mutation);
        assertEquals(2, repeated.path("sheets").size());
    }

    @Test
    void removesTheDetailRegionSourceAndSheetAsOneInverseMutation() throws Exception {
        ObjectNode snapshot = snapshot();
        JsonNode added = descriptor.apply(snapshot, mutation(params("[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1}]")));
        ((ObjectNode) added.path("sheets").get(1)).put("name", "Renamed details");
        ObjectNode removeParams = mapper.createObjectNode().put("targetSheetId", "detail")
                .put("sourceId", "detail-source").put("regionId", "detail-region");
        JsonNode removed = new PivotDrillDownMutationDescriptor("pivot.drilldown.remove")
                .apply(added, new OperationMutation("pivot.drilldown.remove", "detail", removeParams));
        assertEquals(1, removed.path("sheets").size());
        assertEquals(0, removed.path("dataModel").path("sources").size());
    }

    @Test
    void rejectsInverseRemovalWhileAnotherPivotReferencesTheDetailSource() throws Exception {
        ObjectNode snapshot = snapshot();
        JsonNode added = descriptor.apply(snapshot, mutation(params("[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1}]")));
        ((com.fasterxml.jackson.databind.node.ArrayNode) added.path("sheets").get(0).path("pivots"))
                .addObject().putObject("source").put("kind", "data-source").put("dataSourceId", "detail-source");
        ObjectNode removeParams = mapper.createObjectNode().put("targetSheetId", "detail")
                .put("sourceId", "detail-source").put("regionId", "detail-region");
        PivotDrillDownMutationDescriptor remove = new PivotDrillDownMutationDescriptor("pivot.drilldown.remove");
        assertThrows(ServiceException.class, () -> remove.apply(added, new OperationMutation("pivot.drilldown.remove", "detail", removeParams)));
        assertEquals(2, added.path("sheets").size());
        assertEquals(1, added.path("dataModel").path("sources").size());
    }

    @Test
    void rejectsMaterializedMetadataThatDoesNotMatchSourceColumns() throws Exception {
        ObjectNode snapshot = snapshot();
        ObjectNode params = params("[{\"sourceId\":\"orders\",\"recordId\":\"a\",\"sheetId\":\"sheet-1\",\"row\":1}]");
        ((ObjectNode) params.path("detail").path("source").path("fields").get(0)).put("name", "Forged");
        assertThrows(ServiceException.class, () -> descriptor.apply(snapshot, mutation(params)));
        assertEquals(1, snapshot.path("sheets").size());
        assertEquals(0, snapshot.path("dataModel").path("sources").size());
    }

    private OperationMutation mutation(ObjectNode params) {
        return new OperationMutation("pivot.drilldown.add", "sheet-1", params);
    }

    private ObjectNode params(String paths) throws Exception {
        ObjectNode params = mapper.createObjectNode().put("sheetId", "sheet-1").put("pivotId", "pivot")
                .put("label", "Details").put("targetSheetId", "detail");
        JsonNode sourcePaths = mapper.readTree(paths);
        params.set("sourceRowPaths", sourcePaths);
        params.putObject("target").put("row", 0).put("column", 0);
        java.util.Set<String> records = new java.util.LinkedHashSet<>();
        for (JsonNode path : sourcePaths) records.add(path.has("recordId") ? path.path("recordId").asText() : path.path("sheetId").asText() + ":" + path.path("row").asText());
        int rowCount = records.size();
        ObjectNode detail = params.putObject("detail");
        ObjectNode source = detail.putObject("source");
        source.put("schema", "DataSourceManifest").put("version", 1).put("id", "detail-source").put("name", "Details")
                .put("kind", "chunked-table").put("sourceSheetId", "detail").put("rowCount", rowCount).put("blockRowCount", 65536).put("revision", 0);
        source.putObject("sourceRange").put("sheetId", "detail").put("startRow", 0).put("endRow", rowCount).put("startColumn", 0).put("endColumn", 3);
        var fields = source.putArray("fields");
        String[] headers = {"CustomerId", "Amount", "Sheet1.CustomerId", "Region"};
        for (int index = 0; index < headers.length; index++) fields.addObject().put("id", "detail-field-" + index).put("name", headers[index]).put("ordinal", index).put("type", "mixed");
        source.putArray("blocks").addObject().put("id", "detail-block").put("dataSourceId", "detail-source")
                .put("startRow", 0).put("rowCount", rowCount).put("storageKey", "detail-block").put("checksum", "b".repeat(64))
                .put("byteLength", 1).put("encoding", "columnar-v1").put("revision", 0);
        ObjectNode region = detail.putObject("region");
        region.put("id", "detail-region").put("sourceId", "detail-source").put("headerRow", 0).put("revision", 0);
        region.putObject("range").put("sheetId", "detail").put("startRow", 0).put("endRow", rowCount).put("startColumn", 0).put("endColumn", 3);
        var detailHeaders = detail.putArray("headers");
        for (String header : headers) detailHeaders.add(header);
        return params;
    }

    private ObjectNode snapshot() throws Exception {
        return (ObjectNode) mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,
                  "cells":{"0":{"0":{"value":"CustomerId"},"1":{"value":"Amount"},"4":{"value":"CustomerId"},"5":{"value":"Region"}},
                    "1":{"0":{"value":"c1"},"1":{"value":100},"4":{"value":"c2"},"5":{"value":"West"}},
                    "2":{"0":{"value":"c2"},"1":{"value":200},"4":{"value":"c1"},"5":{"value":"East"}}},
                  "pivots":[{"schema":"PivotDefinition","id":"pivot",
                    "source":{"kind":"worksheet-ranges","ranges":[
                      {"sourceId":"orders","range":{"sheetId":"sheet-1","startRow":0,"endRow":2,"startColumn":0,"endColumn":1}},
                      {"sourceId":"customers","range":{"sheetId":"sheet-1","startRow":0,"endRow":2,"startColumn":4,"endColumn":5}}],
                      "relationships":[{"id":"orders-customers","left":{"sourceId":"orders","fieldId":"source:orders:column:0"},
                        "right":{"sourceId":"customers","fieldId":"source:customers:column:0"},"join":"left"}]},
                    "target":{"sheetId":"sheet-1","anchor":{"row":8,"column":0}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[]},
                    "layout":{"rows":[],"columns":[],"filters":[],"values":[],"allowMultipleFiltersPerField":true,
                      "collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},
                      "subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},
                    "refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}]}]}
                """);
    }
}
