package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReportSheetStructuralTransformTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void rowInsertionMapsBindingAnchorsAndRepeatedHeaderRows() {
        ObjectNode snapshot = snapshotWithReportBinding(4);

        StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", "rows.inserted", FormulaReferenceTransformer.Axis.ROW,
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

        assertThrows(ServiceException.class, () -> StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", "rows.deleted",
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
                "sheet-1", "columns.inserted", FormulaReferenceTransformer.Axis.COLUMN, 2, 1, FormulaReferenceTransformer.Direction.INSERT));

        assertEquals("SERVICE_UNAVAILABLE", error.code());
        assertEquals(before, snapshot);
    }

    @Test
    void insertingColumnBeforeSheetTableShiftsRangeWithoutChangingColumnSchema() {
        ObjectNode snapshot = snapshotWithSheetTable();

        StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", "columns.inserted", FormulaReferenceTransformer.Axis.COLUMN,
                1, 1, FormulaReferenceTransformer.Direction.INSERT);
        StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", "columns.inserted", FormulaReferenceTransformer.Axis.COLUMN,
                4, 1, FormulaReferenceTransformer.Direction.INSERT);

        JsonNode table = snapshot.path("sheets").get(0).path("sheetTables").get(0);
        assertEquals(2, table.path("range").path("startColumn").asInt());
        assertEquals(3, table.path("range").path("endColumn").asInt());
        assertEquals(2, table.path("columns").size());
    }

    @Test
    void deletingReferencedRowReturnsServerOwnedFormulaOwnerDelta() {
        ObjectNode snapshot = snapshotWithReportBinding(10);
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        ((ObjectNode) sheet.path("cells")).putObject("0").putObject("1").put("formula", "=A5");

        var patch = StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", "rows.deleted",
                FormulaReferenceTransformer.Axis.ROW, 4, 1, FormulaReferenceTransformer.Direction.DELETE);

        assertEquals(StructuralPatch.VERSION, patch.version());
        assertEquals("rows.deleted", patch.mutationId());
        assertFalse(patch.formulaOwnerDeltas().isEmpty());
        var delta = patch.formulaOwnerDeltas().get(0);
        assertEquals(new StructuralPatch.CellAddress("sheet-1", 0, 1), delta.beforeAddress());
        assertEquals(new StructuralPatch.CellAddress("sheet-1", 0, 1), delta.afterAddress());
        assertEquals("=A5", delta.before().formula());
        assertEquals("=#REF!", delta.after().formula());
        assertEquals("=#REF!", snapshot.path("sheets").get(0).path("cells").path("0").path("1").path("formula").asText());
    }

    @Test
    void inverseStructuralPatchRestoresFormulaAndRejectsChangedOwner() {
        ObjectNode snapshot = snapshotWithReportBinding(10);
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        ((ObjectNode) sheet.path("cells")).putObject("0").putObject("1").put("formula", "=A5");

        var deleted = StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", "rows.deleted",
                FormulaReferenceTransformer.Axis.ROW, 4, 1, FormulaReferenceTransformer.Direction.DELETE);
        StructuralSnapshotReducer.applyAxis(snapshot, "sheet-1", "rows.inserted",
                FormulaReferenceTransformer.Axis.ROW, 4, 1, FormulaReferenceTransformer.Direction.INSERT);

        ObjectNode restored = (ObjectNode) StructuralSnapshotReducer.applyStructuralOwnerPatch(
                snapshot, deleted.inverse("rows.inserted"));
        assertEquals("=A5", restored.path("sheets").get(0).path("cells").path("0").path("1").path("formula").asText());

        ObjectNode changed = restored.deepCopy();
        ((ObjectNode) changed.path("sheets").get(0).path("cells").path("0").path("1")).put("formula", "=B1");
        assertThrows(ServiceException.class, () -> StructuralSnapshotReducer.applyStructuralOwnerPatch(
                changed, deleted.inverse("rows.inserted")));
    }

    @Test
    void cellShiftCollectsFormulaOwnerDeltasOutsideTheShiftedCells() {
        ObjectNode snapshot = snapshotWithReportBinding(10);
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        ((ObjectNode) sheet.path("cells")).putObject("0").putObject("1").put("formula", "=A5");
        RangeRef selection = new RangeRef("sheet-1", 4, 4, 0, 0);
        RangeRef affectedBand = new RangeRef("sheet-1", 4, 19, 0, 0);

        var patch = StructuralSnapshotReducer.shiftCells(snapshot, "sheet-1", "cells.deleted",
                selection, "delete", "row", affectedBand);

        assertEquals("cells.deleted", patch.mutationId());
        assertFalse(patch.formulaOwnerDeltas().isEmpty());
        assertTrue(patch.rangeOwnerDeltas().isEmpty());
        var delta = patch.formulaOwnerDeltas().get(0);
        assertEquals(new StructuralPatch.CellAddress("sheet-1", 0, 1), delta.beforeAddress());
        assertEquals(new StructuralPatch.CellAddress("sheet-1", 0, 1), delta.afterAddress());
        assertEquals("=A5", delta.before().formula());
        assertEquals("=#REF!", delta.after().formula());
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
