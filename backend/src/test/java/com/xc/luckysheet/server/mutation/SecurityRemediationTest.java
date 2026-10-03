package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.config.QuerySource;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.WorkbookSnapshotValidator;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;
import java.util.Map;
import java.util.Set;
import static org.junit.jupiter.api.Assertions.*;

class SecurityRemediationTest {
    private final ObjectMapper mapper = new ObjectMapper();
    @Test void chartElementsPreserveOptionalFieldsAndCanonicalVisibilityModes() throws Exception {
        var workbook = mapper.readTree("{\"sheets\":[{\"id\":\"s\",\"rowCount\":10,\"columnCount\":10}]}");
        var chart = (ObjectNode) mapper.readTree("{\"kind\":\"chart\",\"chartId\":\"c\",\"chartType\":\"column\",\"subtype\":\"clustered\",\"elements\":{\"hiddenData\":\"hideRows\"},\"source\":{\"kind\":\"worksheet-ranges\",\"ranges\":[{\"sheetId\":\"s\",\"startRow\":0,\"endRow\":2,\"startColumn\":0,\"endColumn\":1}]}}");
        for (String mode : java.util.List.of("show", "hideRows", "hideColumns")) {
            ((ObjectNode) chart.get("elements")).put("hiddenData", mode);
            assertDoesNotThrow(() -> WorkbookSnapshotValidator.requireDrawingPayload(workbook, chart));
        }

        ObjectNode mapChart = chart.deepCopy().put("chartType", "map").put("subtype", "filled-map");
        mapChart.putObject("mapOptions").putObject("resource").putArray("features");
        assertDoesNotThrow(() -> WorkbookSnapshotValidator.requireDrawingPayload(workbook, mapChart));
        var feature = ((com.fasterxml.jackson.databind.node.ArrayNode) mapChart.path("mapOptions").path("resource").path("features")).addObject();
        feature.putArray("polygons").addArray().addArray().add(181).add(91);
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireDrawingPayload(workbook, mapChart));
        ((ObjectNode) chart.get("elements")).put("hiddenData", "hide");
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireDrawingPayload(workbook, chart));
    }
    @Test void numericFillRecognizesQuotedFormatsAndPlansLargeTracksOnce() {
        int seeds = 5000;
        ObjectNode root = mapper.createObjectNode();
        ObjectNode sheet = root.putArray("sheets").addObject().put("id", "s").put("rowCount", seeds * 2).put("columnCount", 1);
        ObjectNode cells = sheet.putObject("cells");
        for (int i = 0; i < seeds; i++) cells.putObject(Integer.toString(i)).putObject("0").put("value", i).put("numberFormat", "\"USD\" 0");
        ObjectNode params = mapper.createObjectNode().put("sheetId", "s").put("direction", "down").put("mode", "series");
        params.putObject("series").put("type", "linear").put("trend", true);
        params.putObject("sourceRange").put("sheetId", "s").put("startRow", 0).put("endRow", seeds - 1).put("startColumn", 0).put("endColumn", 0);
        params.putObject("targetRange").put("sheetId", "s").put("startRow", 0).put("endRow", seeds * 2 - 1).put("startColumn", 0).put("endColumn", 0);
        var writes = params.putArray("writes");
        for (int i = seeds; i < seeds * 2; i++) writes.addObject().put("row", i).put("column", 0).putNull("before").putObject("after").put("value", i).put("numberFormat", "\"USD\" 0");
        var mutation = new OperationMutation("fill.applied", "s", params);
        var result = assertTimeout(java.time.Duration.ofSeconds(3), () -> new FillMutationDescriptor("fill.applied").apply(root, mutation));
        assertEquals(seeds * 2 - 1, result.path("sheets").get(0).path("cells").path(Integer.toString(seeds * 2 - 1)).path("0").path("value").asDouble());
        assertFalse(cells.has(Integer.toString(seeds)));
    }
    @Test void configuredSourcesRequireExplicitWorkbookAndSubjectGrants() {
        QuerySource source = new QuerySource("sqlite", "jdbc:sqlite:file.db", null, null, null, Map.of(), Set.of("book"), Set.of("editor"));
        assertDoesNotThrow(() -> source.requireAccess("book", "editor"));
        assertThrows(ServiceException.class, () -> source.requireAccess("other", "editor"));
        assertThrows(ServiceException.class, () -> source.requireAccess("book", "other"));
        assertThrows(ServiceException.class, () -> new QuerySource("sqlite", "jdbc:sqlite:file.db", null, null, null, Map.of(), Set.of(), Set.of()).requireAccess("book", "editor"));
    }
    @Test void payloadBudgetsRejectOversizedTextAndNestedGeometryBeforeRendering() throws Exception {
        assertDoesNotThrow(() -> WorkbookSnapshotValidator.requireResourceBudget(mapper.readTree("{\"text\":\"hello\",\"width\":480,\"fontFamily\":\"Calibri\"}")));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireResourceBudget(mapper.createObjectNode().put("text", "x".repeat(32768))));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireResourceBudget(mapper.readTree("{\"rowHeightsPx\":{\"1\":1000000}}")));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireResourceBudget(mapper.readTree("{\"numberFormat\":{\"invalid\":true}}")));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireResourceBudget(mapper.readTree("{\"fontFamily\":\"Calibri\\nInjected\"}")));
    }
    @Test void tableSheetAndImageContractsAcceptNormalDataAndRejectStoredCrashPayloads() throws Exception {
        assertDoesNotThrow(() -> WorkbookSnapshotValidator.requireTableSheetDefinition(mapper.readTree("{\"viewId\":\"v\",\"columns\":[{\"fieldId\":\"a\",\"caption\":\"A\"}],\"grouping\":[]}")));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireTableSheetDefinition(mapper.readTree("{\"viewId\":\"v\"}")));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireTableSheetDefinition(mapper.readTree("{\"viewId\":\"v\",\"columns\":[],\"grouping\":[{\"fieldId\":\"missing\"}]}")));
        assertDoesNotThrow(() -> WorkbookSnapshotValidator.requireDrawingPayload(mapper.createObjectNode(), mapper.readTree("{\"kind\":\"image\",\"crop\":{\"left\":0.1,\"right\":0,\"top\":0,\"bottom\":0},\"effects\":{\"brightness\":0.2}}")));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireDrawingPayload(mapper.createObjectNode(), mapper.readTree("{\"kind\":\"image\",\"crop\":{\"left\":0.7,\"right\":0.7,\"top\":0,\"bottom\":0}}")));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireDrawingPayload(mapper.createObjectNode(), mapper.readTree("{\"kind\":\"image\",\"effects\":{\"transparency\":-1}}")));
    }
    @Test void checkboxStatesAreDistinctAndNormalizeBooleanAliases() throws Exception {
        var editor = mapper.readTree("{\"kind\":\"checkbox\",\"trueValue\":\"yes\",\"falseValue\":\"no\"}");
        assertEquals("yes", WorkbookSnapshotValidator.normalizeCheckboxValue(editor, mapper.readTree("1")).asText());
        assertEquals("no", WorkbookSnapshotValidator.normalizeCheckboxValue(editor, mapper.readTree("\"FALSE\"")).asText());
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.requireCellEditor(mapper.readTree("{\"kind\":\"checkbox\",\"trueValue\":1,\"falseValue\":1}")));
        assertThrows(ServiceException.class, () -> WorkbookSnapshotValidator.normalizeCheckboxValue(editor, mapper.readTree("2")));
    }
    @Test void fillPlanRejectsForgedValuesAndMetadataWithoutMutatingTheInput() throws Exception {
        var snapshot = mapper.readTree("{\"sheets\":[{\"id\":\"s\",\"rowCount\":10,\"columnCount\":10,\"cells\":{\"0\":{\"0\":{\"value\":1}}}}]}");
        ObjectNode params = (ObjectNode) mapper.readTree("{\"sheetId\":\"s\",\"sourceRange\":{\"sheetId\":\"s\",\"startRow\":0,\"endRow\":0,\"startColumn\":0,\"endColumn\":0},\"targetRange\":{\"sheetId\":\"s\",\"startRow\":0,\"endRow\":1,\"startColumn\":0,\"endColumn\":0},\"direction\":\"down\",\"mode\":\"copy\",\"writes\":[{\"row\":1,\"column\":0,\"before\":null,\"after\":{\"value\":1}}]}");
        var reducer = new FillMutationDescriptor("fill.applied");
        assertEquals(1, reducer.apply(snapshot, new OperationMutation("fill.applied", "s", params)).path("sheets").get(0).path("cells").path("1").path("0").path("value").asInt());
        ((ObjectNode) params.path("writes").get(0).path("after")).put("value", 999);
        assertThrows(ServiceException.class, () -> reducer.apply(snapshot, new OperationMutation("fill.applied", "s", params)));
        ((ObjectNode) params.path("writes").get(0).path("after")).put("value", 1).putObject("editor").put("kind", "custom").put("editorId", "forged");
        assertThrows(ServiceException.class, () -> reducer.apply(snapshot, new OperationMutation("fill.applied", "s", params)));
        assertFalse(snapshot.path("sheets").get(0).path("cells").has("1"));
    }
}
