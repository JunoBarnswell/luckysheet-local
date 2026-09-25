package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
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

    @Test
    void insertingColumnInsideSheetTableRejectsBeforeChangingSnapshot() {
        ObjectNode snapshot = snapshotWithSheetTable();
        JsonNode before = snapshot.deepCopy();

        ServiceException error = assertThrows(ServiceException.class, () -> StructuralSnapshotReducer.applyAxis(snapshot,
                "sheet-1", FormulaReferenceTransformer.Axis.COLUMN, 2, 1, FormulaReferenceTransformer.Direction.INSERT));

        assertEquals("SERVICE_UNAVAILABLE", error.code());
        assertEquals(before, snapshot);
    }

    @Test
    void insertingColumnBeforeSheetTableShiftsRangeWithoutChangingColumnSchema() {
        ObjectNode snapshot = snapshotWithSheetTable();

        StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", FormulaReferenceTransformer.Axis.COLUMN,
                1, 1, FormulaReferenceTransformer.Direction.INSERT);
        StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", FormulaReferenceTransformer.Axis.COLUMN,
                4, 1, FormulaReferenceTransformer.Direction.INSERT);

        JsonNode table = snapshot.path("sheets").get(0).path("sheetTables").get(0);
        assertEquals(2, table.path("range").path("startColumn").asInt());
        assertEquals(3, table.path("range").path("endColumn").asInt());
        assertEquals(2, table.path("columns").size());
    }

    private ObjectNode snapshotWithSheetTable() {
        ObjectNode snapshot = snapshotWithReportBinding(4);
        ObjectNode table = ((ObjectNode) snapshot.path("sheets").get(0)).withArray("sheetTables").addObject();
        table.put("id", "table-1").put("sheetId", "sheet-1").put("name", "Table1");
        table.putObject("range").put("sheetId", "sheet-1").put("startRow", 0).put("endRow", 2)
                .put("startColumn", 1).put("endColumn", 2);
        table.put("hasHeaderRow", true).put("hasTotalRow", false).put("showBandedRows", true)
                .put("showBandedColumns", false).put("showFirstColumn", false).put("showLastColumn", false)
                .put("showFilterButton", false).put("autoExpand", "none");
        ArrayNode columns = table.putArray("columns");
        columns.addObject().put("id", "column-1").put("name", "A");
        columns.addObject().put("id", "column-2").put("name", "B");
        return snapshot;
    }

    private ObjectNode snapshotWithReportBinding(int bindingRow) {
        ObjectNode snapshot = mapper.createObjectNode();
        ObjectNode sheet = snapshot.putArray("sheets").addObject();
        sheet.put("id", "sheet-1").put("name", "Report").put("kind", "report-sheet").put("rowCount", 20).put("columnCount", 10);
        sheet.putObject("pane").put("kind", "none");
        sheet.putObject("cells");
        sheet.putArray("sheetTables");
        ObjectNode review = sheet.putObject("review");
        review.putObject("notesByCell");
        review.putObject("notesById");
        review.putObject("threadIdsByCell");
        review.putObject("threadsById");
        ObjectNode report = sheet.putObject("reportSheet");
        ObjectNode binding = report.putArray("bindings").addObject();
        binding.putObject("cell").put("row", bindingRow).put("column", 2);
        binding.put("expression", "field-id").put("kind", "field");
        report.putObject("pagination").put("enabled", true).putArray("repeatHeaderRows").add(0).add(bindingRow);
        return snapshot;
    }
}
