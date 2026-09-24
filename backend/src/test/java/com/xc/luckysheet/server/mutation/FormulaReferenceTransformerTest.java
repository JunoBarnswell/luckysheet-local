package com.xc.luckysheet.server.mutation;

import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.time.Duration;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTimeout;

class FormulaReferenceTransformerTest {
    private final FormulaReferenceTransformer.SheetIdentity sheet = new FormulaReferenceTransformer.SheetIdentity("sheet-1", "Sheet1");

    @Test
    void axisTransformMovesAbsoluteAndRelativeReferencesWithoutTouchingStringsOrTables() {
        String formula = "=SUM(A1,$B$2,Table[Amount],\"A1\")";
        String result = FormulaReferenceTransformer.remapAxis(
                formula,
                sheet,
                sheet,
                FormulaReferenceTransformer.Axis.ROW,
                1,
                2,
                FormulaReferenceTransformer.Direction.INSERT
        );

        assertEquals("=SUM(A1,$B$4,Table[Amount],\"A1\")", result);
    }

    @Test
    void deletionCreatesRealReferenceErrorForDeletedRangeEndpoint() {
        String result = FormulaReferenceTransformer.remapAxis(
                "=A2:B3",
                sheet,
                sheet,
                FormulaReferenceTransformer.Axis.ROW,
                1,
                1,
                FormulaReferenceTransformer.Direction.DELETE
        );

        assertEquals("=#REF!", result);
    }

    @Test
    void renameChangesOnlyQualifiedSheetReferenceSyntax() {
        String result = FormulaReferenceTransformer.renameSheet("='Old Name'!A1+Old!B2+\"Old!C3\"", "Old Name", "New Name");

        assertEquals("='New Name'!A1+Old!B2+\"Old!C3\"", result);
    }

    @Test
    void moveOffsetHonorsAbsoluteMarkers() {
        assertEquals("=C3+$B$1+E$1+$D3", FormulaReferenceTransformer.offset("=A1+$B$1+C$1+$D1", 2, 2));
    }

    @Test
    void removedQualifiedSheetReferenceBecomesReferenceError() {
        assertEquals("=#REF!+A1", FormulaReferenceTransformer.invalidateSheet("=Sheet1!A1+A1", "sheet-1", "Sheet1"));
    }

    @Test
    void rangeUnionIntersectionAndDynamicArraySyntaxRemainStructuredFormulaText() {
        String result = FormulaReferenceTransformer.remapAxis(
                "=SUM(Sheet1!A2:B3,A1 B1,@C2#)",
                sheet,
                sheet,
                FormulaReferenceTransformer.Axis.ROW,
                1,
                1,
                FormulaReferenceTransformer.Direction.INSERT
        );

        assertEquals("=SUM(Sheet1!A3:B4,A1 B1,@C3#)", result);
    }

    @Test
    void structuralEditsPreserveThreeDimensionalReferencesWhenTargetIsOutsideTheirSheetSpan() {
        List<FormulaReferenceTransformer.SheetIdentity> order = List.of(
                new FormulaReferenceTransformer.SheetIdentity("sheet-1", "Sheet1"),
                new FormulaReferenceTransformer.SheetIdentity("sheet-2", "Sheet2"),
                new FormulaReferenceTransformer.SheetIdentity("sheet-3", "Sheet3"),
                new FormulaReferenceTransformer.SheetIdentity("sheet-4", "Sheet4"));
        FormulaReferenceTransformer.SheetIdentity target = order.get(3);
        String formula = "=SUM(Sheet1:Sheet3!A1)+A1";

        assertEquals("=SUM(Sheet1:Sheet3!A1)+A2", FormulaReferenceTransformer.remapAxis(
                formula, target, target, FormulaReferenceTransformer.Axis.ROW, 0, 1,
                FormulaReferenceTransformer.Direction.INSERT, order));
        assertEquals("=SUM(Sheet1:Sheet3!A1)+A2", FormulaReferenceTransformer.remapCellShift(
                formula, target, target, new FormulaReferenceTransformer.Range(0, 0, 0, 0),
                FormulaReferenceTransformer.Axis.ROW, FormulaReferenceTransformer.Direction.INSERT, order));
    }

    @Test
    void structuralEditsRejectThreeDimensionalReferencesWhenTargetIsInsideTheirSheetSpan() {
        List<FormulaReferenceTransformer.SheetIdentity> order = List.of(
                new FormulaReferenceTransformer.SheetIdentity("sheet-1", "Sheet1"),
                new FormulaReferenceTransformer.SheetIdentity("sheet-2", "Sheet2"),
                new FormulaReferenceTransformer.SheetIdentity("sheet-3", "Sheet3"));
        FormulaReferenceTransformer.SheetIdentity target = order.get(1);
        String formula = "=SUM(Sheet1:Sheet3!A1)";

        assertThrows(ServiceException.class, () -> FormulaReferenceTransformer.remapAxis(
                formula, target, target, FormulaReferenceTransformer.Axis.ROW, 0, 1,
                FormulaReferenceTransformer.Direction.INSERT, order));
        assertThrows(ServiceException.class, () -> FormulaReferenceTransformer.remapCellShift(
                formula, target, target, new FormulaReferenceTransformer.Range(0, 0, 0, 0),
                FormulaReferenceTransformer.Axis.ROW, FormulaReferenceTransformer.Direction.INSERT, order));
    }

    @Test
    void renameUpdatesOnlyMatchingEndpointsOfThreeDimensionalReferences() {
        assertEquals("=SUM('New Name:Sheet3'!A1)+Old!B2",
                FormulaReferenceTransformer.renameSheet("=SUM('Old Name:Sheet3'!A1)+Old!B2", "Old Name", "New Name"));
        assertEquals("=SUM(Sheet1:'New Name'!A1)",
                FormulaReferenceTransformer.renameSheet("=SUM(Sheet1:'Old Name'!A1)", "Old Name", "New Name"));
    }

    @Test
    void longNonReferenceIdentifierIsScannedInLinearTime() {
        String formula = "=" + "A".repeat(128_000);

        String result = assertTimeout(
                Duration.ofSeconds(2),
                () -> FormulaReferenceTransformer.renameSheet(formula, "Old", "New")
        );

        assertEquals(formula, result);
    }
}
