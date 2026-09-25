package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class ReportSheetStructuralTransformTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void rowInsertionMapsBindingAnchorsAndRepeatedHeaderRows() {
        ObjectNode snapshot = snapshotWithReportBinding(4);

        StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", FormulaReferenceTransformer.Axis.ROW,
                2, 2, FormulaReferenceTransformer.Direction.INSERT);

        JsonNode report = snapshot.path("sheets").get(0).path("reportSheet");
        assertEquals(6, report.path("bindings").get(0).path("cell").path("row").asInt());
        assertEquals(0, report.path("pagination").path("repeatHeaderRows").get(0).asInt());
        assertEquals(6, report.path("pagination").path("repeatHeaderRows").get(1).asInt());
    }

    @Test
    void deletingReportBindingAnchorRejectsBeforeChangingCellsOrReportDefinition() {
        ObjectNode snapshot = snapshotWithReportBinding(4);
        JsonNode sheet = snapshot.path("sheets").get(0);
        JsonNode beforeReport = sheet.path("reportSheet").deepCopy();
        JsonNode beforeCells = sheet.path("cells").deepCopy();

        assertThrows(ServiceException.class, () -> StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1",
                FormulaReferenceTransformer.Axis.ROW, 4, 1, FormulaReferenceTransformer.Direction.DELETE));

        assertEquals(beforeReport, sheet.path("reportSheet"));
        assertEquals(beforeCells, sheet.path("cells"));
        assertEquals(20, sheet.path("rowCount").asInt());
    }

    private ObjectNode snapshotWithReportBinding(int bindingRow) {
        ObjectNode snapshot = mapper.createObjectNode();
        ObjectNode sheet = snapshot.putArray("sheets").addObject();
        sheet.put("id", "sheet-1").put("rowCount", 20).put("columnCount", 10);
        sheet.putObject("pane").put("kind", "none");
        sheet.putObject("cells");
        sheet.putArray("sheetTables");
        ObjectNode report = sheet.putObject("reportSheet");
        ObjectNode binding = report.putArray("bindings").addObject();
        binding.putObject("cell").put("row", bindingRow).put("column", 2);
        binding.put("expression", "field-id").put("kind", "field");
        report.putObject("pagination").put("enabled", true).putArray("repeatHeaderRows").add(0).add(bindingRow);
        return snapshot;
    }
}
