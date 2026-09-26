package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class StructuralPatchMergeTest {
    @Test
    void mergeDeduplicatesEveryCanonicalOwnerKindAndPreservesOrder() {
        StructuralPatch.FormulaOwnerDelta cell = cellDelta("sheet-1", 0, 0, "=A1", "=A2");
        StructuralPatch.FormulaOwnerDelta rule = StructuralPatch.FormulaOwnerDelta.formulaRule(
                "sheet-1", "conditional-format", "cf-1", "value1", "=A1>0", "=A2>0",
                List.of(range("sheet-1", 0, 3, 0, 0)), List.of(range("sheet-1", 0, 3, 0, 0)));
        StructuralPatch.FormulaOwnerDelta chartTitle = StructuralPatch.FormulaOwnerDelta.formulaObject(
                "sheet-1", "chart-text", "chart-1", "titleText.linkedFormula", "=A1", "=A2");
        StructuralPatch.FormulaOwnerDelta shape = objectDelta("shape-property", "sheet-1", "shape-1", null, null, null, null);
        StructuralPatch.FormulaOwnerDelta tableColumn = objectDelta("table-sheet-column", "sheet-1", null, "column-1", null, null, null);
        StructuralPatch.FormulaOwnerDelta dataViewField = objectDelta("data-view-field", null, null, "field-1", "view-1", null, null);
        StructuralPatch.FormulaOwnerDelta styleTemplate = objectDelta("cell-style-template", null, null, null, null, "template-1", "formula1");
        StructuralPatch.FormulaOwnerDelta chartLegend = StructuralPatch.FormulaOwnerDelta.formulaObject(
                "sheet-1", "chart-text", "chart-1", "legend.text.linkedFormula", "=A1", "=A2");
        StructuralPatch.FormulaOwnerDelta nextCell = cellDelta("sheet-1", 0, 1, "=B1", "=B2");

        StructuralPatch.DefinedNameOwnerDelta workbookName = nameDelta("workbook", "Revenue", null, "=A1", "=A2");
        StructuralPatch.DefinedNameOwnerDelta scopedName = nameDelta("sheet", "Revenue", "sheet-1", "=A1", "=A2");
        StructuralPatch generated = patch("rows.inserted",
                List.of(cell, rule, chartTitle, shape, tableColumn, dataViewField, styleTemplate), List.of(workbookName));
        StructuralPatch inverse = patch("rows.inserted",
                List.of(cell, rule, chartTitle, shape, tableColumn, dataViewField, styleTemplate, chartLegend, nextCell),
                List.of(workbookName, scopedName));

        StructuralPatch merged = MutationDescriptorRegistry.mergeStructuralPatches("rows.inserted", generated, inverse);

        assertEquals(List.of(cell, rule, chartTitle, shape, tableColumn, dataViewField, styleTemplate, chartLegend, nextCell),
                merged.formulaOwnerDeltas());
        assertEquals(List.of(workbookName, scopedName), merged.definedNameOwnerDeltas());
    }

    @Test
    void mergeRejectsConflictingFormulaOwnersAtTheSameCanonicalAddress() {
        StructuralPatch.FormulaOwnerDelta generatedDelta = cellDelta("sheet-1", 1, 2, "=A1", "=A2");
        StructuralPatch.FormulaOwnerDelta inverseDelta = cellDelta("sheet-1", 1, 2, "=A1", "=A3");

        ServiceException error = assertThrows(ServiceException.class, () -> MutationDescriptorRegistry.mergeStructuralPatches(
                "rows.inserted", patch("rows.inserted", List.of(generatedDelta), List.of()),
                patch("rows.inserted", List.of(inverseDelta), List.of())));

        assertEquals("CONFLICT", error.code());
    }

    @Test
    void mergeIndexesTheSmallerInverseSideAndAppendsItsUnmatchedOwnersInOrder() {
        StructuralPatch.FormulaOwnerDelta first = cellDelta("sheet-1", 0, 0, "=A1", "=A2");
        StructuralPatch.FormulaOwnerDelta second = cellDelta("sheet-1", 0, 1, "=B1", "=B2");
        StructuralPatch.FormulaOwnerDelta third = cellDelta("sheet-1", 0, 2, "=C1", "=C2");
        StructuralPatch.FormulaOwnerDelta appendedFirst = cellDelta("sheet-1", 0, 3, "=D1", "=D2");

        StructuralPatch merged = MutationDescriptorRegistry.mergeStructuralPatches("rows.inserted",
                patch("rows.inserted", List.of(first, second, third), List.of()),
                patch("rows.inserted", List.of(second, appendedFirst), List.of()));

        assertEquals(List.of(first, second, third, appendedFirst), merged.formulaOwnerDeltas());
    }

    @Test
    void mergeDeduplicatesRangeOwnersAndPreservesGeneratedThenInverseOrder() {
        StructuralPatch.RangeOwnerDelta table = StructuralPatch.RangeOwnerDelta.range(
                "workbook-table", "table-1", range("sheet-1", 0, 2, 0, 1), range("sheet-1", 1, 3, 0, 1));
        StructuralPatch.RangeOwnerDelta region = StructuralPatch.RangeOwnerDelta.dataRegion(
                "sheet-1", "region-1", range("sheet-1", 0, 2, 0, 1), 0,
                range("sheet-1", 1, 3, 0, 1), 1);
        StructuralPatch.RangeOwnerDelta source = StructuralPatch.RangeOwnerDelta.range(
                "data-source", "source-1", range("sheet-1", 4, 6, 0, 1), range("sheet-1", 5, 7, 0, 1));
        StructuralPatch.RangeOwnerDelta sheetTable = StructuralPatch.RangeOwnerDelta.sheetTable("sheet-1", "sheet-table-1",
                range("sheet-1", 8, 10, 0, 1), range("sheet-1", 9, 11, 0, 1));

        StructuralPatch merged = MutationDescriptorRegistry.mergeStructuralPatches("rows.inserted",
                patch("rows.inserted", List.of(), List.of(), List.of(table, region)),
                patch("rows.inserted", List.of(), List.of(), List.of(table, region, source, sheetTable)));

        assertEquals(List.of(table, region, source, sheetTable), merged.rangeOwnerDeltas());
    }

    @Test
    void mergeRejectsConflictingRangeOwnerFactsForOneStableIdentity() {
        StructuralPatch.RangeOwnerDelta generated = StructuralPatch.RangeOwnerDelta.range(
                "workbook-table", "table-1", range("sheet-1", 0, 2, 0, 1), range("sheet-1", 1, 3, 0, 1));
        StructuralPatch.RangeOwnerDelta inverse = StructuralPatch.RangeOwnerDelta.range(
                "workbook-table", "table-1", range("sheet-1", 0, 2, 0, 1), range("sheet-1", 2, 4, 0, 1));

        ServiceException error = assertThrows(ServiceException.class, () -> MutationDescriptorRegistry.mergeStructuralPatches(
                "rows.inserted", patch("rows.inserted", List.of(), List.of(), List.of(generated)),
                patch("rows.inserted", List.of(), List.of(), List.of(inverse))));

        assertEquals("CONFLICT", error.code());
    }

    @Test
    void mergeRejectsDefinedNamesThatDifferOnlyByCaseAsOneOwnerConflict() {
        StructuralPatch.DefinedNameOwnerDelta generatedDelta = nameDelta("workbook", "Revenue", null, "=A1", "=A2");
        StructuralPatch.DefinedNameOwnerDelta inverseDelta = nameDelta("workbook", "revenue", null, "=A1", "=A3");

        ServiceException error = assertThrows(ServiceException.class, () -> MutationDescriptorRegistry.mergeStructuralPatches(
                "rows.inserted", patch("rows.inserted", List.of(), List.of(generatedDelta)),
                patch("rows.inserted", List.of(), List.of(inverseDelta))));

        assertEquals("CONFLICT", error.code());
    }

    @Test
    void structuralPatchSerializesRangeOwnerDeltasAsAnExactTopLevelField() {
        ObjectMapper mapper = new ObjectMapper();
        StructuralPatch.RangeOwnerDelta sheetTable = StructuralPatch.RangeOwnerDelta.sheetTable("sheet-1", "sheet-table-1",
                range("sheet-1", 0, 2, 0, 1), range("sheet-1", 1, 3, 0, 1));
        StructuralPatch patch = patch("rows.inserted", List.of(), List.of(), List.of(sheetTable));

        Set<String> fields = new HashSet<>();
        mapper.valueToTree(patch).fieldNames().forEachRemaining(fields::add);

        assertEquals(Set.of("version", "mutationId", "formulaOwnerDeltas", "definedNameOwnerDeltas", "rangeOwnerDeltas"), fields);
        assertEquals("sheet-1", mapper.valueToTree(patch).path("rangeOwnerDeltas").get(0).path("sheetId").asText());
    }

    private static StructuralPatch patch(String mutationId, List<StructuralPatch.FormulaOwnerDelta> formulas,
            List<StructuralPatch.DefinedNameOwnerDelta> names) {
        return patch(mutationId, formulas, names, List.of());
    }

    private static StructuralPatch patch(String mutationId, List<StructuralPatch.FormulaOwnerDelta> formulas,
            List<StructuralPatch.DefinedNameOwnerDelta> names, List<StructuralPatch.RangeOwnerDelta> ranges) {
        return new StructuralPatch(StructuralPatch.VERSION, mutationId, formulas, names, ranges);
    }

    private static StructuralPatch.FormulaOwnerDelta cellDelta(String sheetId, int row, int column,
            String beforeFormula, String afterFormula) {
        StructuralPatch.CellAddress address = new StructuralPatch.CellAddress(sheetId, row, column);
        return new StructuralPatch.FormulaOwnerDelta("formula-cell", address, address,
                new StructuralPatch.FormulaOwnerState(beforeFormula, null, null),
                new StructuralPatch.FormulaOwnerState(afterFormula, null, null));
    }

    private static StructuralPatch.FormulaOwnerDelta objectDelta(String ownerKind, String sheetId,
            String ownerId, String fieldId, String viewId, String templateId, String field) {
        return StructuralPatch.FormulaOwnerDelta.formulaObject(ownerKind, sheetId, ownerId, fieldId,
                viewId, templateId, field, "=A1", "=A2");
    }

    private static StructuralPatch.DefinedNameOwnerDelta nameDelta(String scope, String name, String sheetId,
            String beforeFormula, String afterFormula) {
        StructuralPatch.DefinedNameOwnerIdentity owner = new StructuralPatch.DefinedNameOwnerIdentity(scope, name, sheetId);
        StructuralPatch.DefinedNameState before = new StructuralPatch.DefinedNameState(name, beforeFormula, scope, sheetId, null);
        StructuralPatch.DefinedNameState after = new StructuralPatch.DefinedNameState(name, afterFormula, scope, sheetId, null);
        return new StructuralPatch.DefinedNameOwnerDelta(owner, before, after);
    }

    private static RangeRef range(String sheetId, int startRow, int endRow, int startColumn, int endColumn) {
        return new RangeRef(sheetId, startRow, endRow, startColumn, endColumn);
    }
}
