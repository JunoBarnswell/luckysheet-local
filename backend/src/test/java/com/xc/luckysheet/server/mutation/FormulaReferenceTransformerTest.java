package com.xc.luckysheet.server.mutation;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

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
}
