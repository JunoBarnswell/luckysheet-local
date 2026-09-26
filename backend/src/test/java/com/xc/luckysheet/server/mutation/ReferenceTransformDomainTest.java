package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ReferenceTransformDomainTest {
    @Test
    void mapsPointsAndInclusiveIntervalsWithSharedInsertDeleteSemantics() {
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.MAPPED, 2L),
                ReferenceTransformDomain.mapPoint(2, 3, 2, true, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.MAPPED, 5L),
                ReferenceTransformDomain.mapPoint(3, 3, 2, true, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.DELETED, null),
                ReferenceTransformDomain.mapPoint(3, 3, 2, false, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.MAPPED, 3L),
                ReferenceTransformDomain.mapPoint(5, 3, 2, false, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.IntervalMapping(ReferenceTransformDomain.IntervalKind.MAPPED, 2L, 6L),
                ReferenceTransformDomain.mapInterval(2, 4, 3, 2, true, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.IntervalMapping(ReferenceTransformDomain.IntervalKind.MAPPED, 2L, 3L),
                ReferenceTransformDomain.mapInterval(2, 5, 3, 2, false, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.IntervalMapping(ReferenceTransformDomain.IntervalKind.DELETED, null, null),
                ReferenceTransformDomain.mapInterval(3, 4, 3, 2, false, ReferenceTransformDomain.MAX_ROW_INDEX));
    }

    @Test
    void rejectsIntervalsOutsideTheWorksheetAddressLimit() {
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.OUT_OF_BOUNDS,
                        ReferenceTransformDomain.MAX_COLUMN_INDEX + 1L),
                ReferenceTransformDomain.mapPoint(ReferenceTransformDomain.MAX_COLUMN_INDEX,
                        ReferenceTransformDomain.MAX_COLUMN_INDEX, 1, true, ReferenceTransformDomain.MAX_COLUMN_INDEX));
        assertEquals(new ReferenceTransformDomain.IntervalMapping(ReferenceTransformDomain.IntervalKind.OUT_OF_BOUNDS,
                        ReferenceTransformDomain.MAX_COLUMN_INDEX + 1L, ReferenceTransformDomain.MAX_COLUMN_INDEX + 1L),
                ReferenceTransformDomain.mapInterval(ReferenceTransformDomain.MAX_COLUMN_INDEX,
                        ReferenceTransformDomain.MAX_COLUMN_INDEX, ReferenceTransformDomain.MAX_COLUMN_INDEX, 1, true,
                        ReferenceTransformDomain.MAX_COLUMN_INDEX));
    }

    @Test
    void rejectsStructuralIntervalsOutsideTheWorksheetAddressDomain() {
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.mapInterval(
                0, 0, ReferenceTransformDomain.MAX_ROW_INDEX + 1, 1, true, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.mapInterval(
                0, 0, ReferenceTransformDomain.MAX_ROW_INDEX, 2, false, ReferenceTransformDomain.MAX_ROW_INDEX));
    }

    @Test
    void matchesTheSharedTypeScriptAndJavaStructuralMappingVectors() throws IOException {
        try (InputStream vectorsStream = getClass().getResourceAsStream("/reference-transform-vectors.json")) {
            assertNotNull(vectorsStream, "shared structural transform vectors must be on the test classpath");
            JsonNode vectors = new ObjectMapper().readTree(vectorsStream);
            assertEquals("ReferenceTransformVectors", vectors.path("schema").asText());
            assertEquals(3, vectors.path("version").asInt());

            for (JsonNode vector : vectors.path("points")) {
                JsonNode expected = vector.path("expected");
                ReferenceTransformDomain.PointMapping actual = ReferenceTransformDomain.mapPoint(
                        vector.path("position").asInt(),
                        vector.path("at").asInt(),
                        vector.path("count").asInt(),
                        "insert".equals(vector.path("operation").asText()),
                        maximum(vector));
                Long expectedPosition = expected.has("position") ? expected.path("position").asLong() : null;
                assertEquals(new ReferenceTransformDomain.PointMapping(
                                ReferenceTransformDomain.PointKind.valueOf(expected.path("kind").asText().replace('-', '_').toUpperCase(Locale.ROOT)),
                                expectedPosition),
                        actual, vector.path("id").asText());
            }

            for (JsonNode vector : vectors.path("intervals")) {
                JsonNode expected = vector.path("expected");
                Long expectedStart = expected.has("start") ? expected.path("start").asLong() : null;
                Long expectedEnd = expected.has("end") ? expected.path("end").asLong() : null;
                ReferenceTransformDomain.IntervalMapping actual = ReferenceTransformDomain.mapInterval(
                        vector.path("start").asInt(),
                        vector.path("end").asInt(),
                        vector.path("at").asInt(),
                        vector.path("count").asInt(),
                        "insert".equals(vector.path("operation").asText()),
                        maximum(vector));
                assertEquals(new ReferenceTransformDomain.IntervalMapping(
                                ReferenceTransformDomain.IntervalKind.valueOf(expected.path("kind").asText().replace('-', '_').toUpperCase(Locale.ROOT)),
                                expectedStart,
                                expectedEnd),
                        actual, vector.path("id").asText());
            }

            for (JsonNode vector : vectors.path("cellShiftIndices")) {
                assertEquals(vector.path("expected").asLong(), ReferenceTransformDomain.mapCellShiftIndex(
                        vector.path("position").asInt(),
                        vector.path("start").asInt(),
                        vector.path("end").asInt(),
                        ReferenceTransformDomain.Operation.valueOf(vector.path("operation").asText().toUpperCase(Locale.ROOT)),
                        maximum(vector)),
                        vector.path("id").asText());
            }

            for (JsonNode vector : vectors.path("cellShiftPoints")) {
                JsonNode selection = vector.path("selection");
                JsonNode expected = vector.path("expected");
                Integer expectedRow = expected.has("row") ? expected.path("row").asInt() : null;
                Integer expectedColumn = expected.has("column") ? expected.path("column").asInt() : null;
                ReferenceTransformDomain.CellPointMapping actual = ReferenceTransformDomain.mapCellShiftPoint(
                        vector.path("row").asInt(),
                        vector.path("column").asInt(),
                        selection.path("startRow").asInt(),
                        selection.path("endRow").asInt(),
                        selection.path("startColumn").asInt(),
                        selection.path("endColumn").asInt(),
                        ReferenceTransformDomain.CellAxis.valueOf(vector.path("axis").asText().toUpperCase(Locale.ROOT)),
                        ReferenceTransformDomain.Operation.valueOf(vector.path("operation").asText().toUpperCase(Locale.ROOT)));
                assertEquals(new ReferenceTransformDomain.CellPointMapping(
                                ReferenceTransformDomain.CellPointKind.valueOf(expected.path("kind").asText().replace('-', '_').toUpperCase(Locale.ROOT)),
                                expectedRow,
                                expectedColumn),
                        actual, vector.path("id").asText());
            }

            for (JsonNode vector : vectors.path("rowPermutations")) {
                JsonNode sourceRows = vector.path("sourceRowsByTarget");
                int[] input = new int[sourceRows.size()];
                for (int index = 0; index < sourceRows.size(); index++) input[index] = sourceRows.get(index).asInt();
                int[] expected = new int[vector.path("expectedTargetRowsBySource").size()];
                for (int index = 0; index < expected.length; index++) expected[index] = vector.path("expectedTargetRowsBySource").get(index).asInt();
                assertArrayEquals(expected,
                        ReferenceTransformDomain.createRowPermutationMap(vector.path("startRow").asInt(), input),
                        vector.path("id").asText());
            }
            for (JsonNode vector : vectors.path("permutationPoints")) {
                int[] targetRowsBySource = new int[vector.path("targetRowsBySource").size()];
                for (int index = 0; index < targetRowsBySource.length; index++) {
                    targetRowsBySource[index] = vector.path("targetRowsBySource").get(index).asInt();
                }
                assertEquals(vector.path("expected").asInt(), ReferenceTransformDomain.mapPermutationIndex(
                        vector.path("position").asInt(), vector.path("startRow").asInt(), targetRowsBySource),
                        vector.path("id").asText());
            }
        }
    }

    @Test
    void rejectsInvalidCellShiftDomainInputs() {
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.mapCellShiftIndex(
                -1, 0, 0, ReferenceTransformDomain.Operation.INSERT, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.mapCellShiftIndex(
                0, 1, 0, ReferenceTransformDomain.Operation.DELETE, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.mapCellShiftPoint(
                0, 0, 0, 0, 0, ReferenceTransformDomain.MAX_COLUMN_INDEX + 1,
                ReferenceTransformDomain.CellAxis.COLUMN, ReferenceTransformDomain.Operation.INSERT));
    }

    @Test
    void rejectsInvalidRowPermutationMapsAndPositions() {
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.createRowPermutationMap(4, new int[]{4, 4}));
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.createRowPermutationMap(
                ReferenceTransformDomain.MAX_ROW_INDEX, new int[]{ReferenceTransformDomain.MAX_ROW_INDEX, ReferenceTransformDomain.MAX_ROW_INDEX + 1}));
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.mapPermutationIndex(
                ReferenceTransformDomain.MAX_ROW_INDEX + 1, 0, new int[]{0}));
        assertThrows(IllegalArgumentException.class, () -> ReferenceTransformDomain.mapPermutationIndex(0, 0, new int[]{1}));
    }

    private static int maximum(JsonNode vector) {
        return switch (vector.path("axis").asText()) {
            case "row" -> ReferenceTransformDomain.MAX_ROW_INDEX;
            case "column" -> ReferenceTransformDomain.MAX_COLUMN_INDEX;
            default -> throw new IllegalArgumentException("Shared reference vector axis is invalid");
        };
    }

    @Test
    void matchesTheSharedTypeScriptAndJavaFormulaAxisVectors() throws IOException {
        try (InputStream vectorsStream = getClass().getResourceAsStream("/reference-transform-vectors.json")) {
            assertNotNull(vectorsStream);
            JsonNode vectors = new ObjectMapper().readTree(vectorsStream);
            List<FormulaReferenceTransformer.SheetIdentity> sheetOrder = new ArrayList<>();
            for (JsonNode sheet : vectors.path("formulaSheetOrder")) {
                sheetOrder.add(new FormulaReferenceTransformer.SheetIdentity(sheet.path("id").asText(), sheet.path("name").asText()));
            }
            assertTrue(vectors.path("formulaAxes").isArray() && !vectors.path("formulaAxes").isEmpty());
            for (JsonNode vector : vectors.path("formulaAxes")) {
                FormulaReferenceTransformer.SheetIdentity owner = sheetOrder.stream()
                        .filter(sheet -> sheet.id().equals(vector.path("ownerSheetId").asText())).findFirst().orElseThrow();
                FormulaReferenceTransformer.SheetIdentity target = sheetOrder.stream()
                        .filter(sheet -> sheet.id().equals(vector.path("targetSheetId").asText())).findFirst().orElseThrow();
                String actual = FormulaReferenceTransformer.remapAxis(
                        vector.path("formula").asText(), owner, target,
                        FormulaReferenceTransformer.Axis.valueOf(vector.path("axis").asText().toUpperCase(Locale.ROOT)),
                        vector.path("at").asInt(), vector.path("count").asInt(),
                        FormulaReferenceTransformer.Direction.valueOf(vector.path("operation").asText().toUpperCase(Locale.ROOT)),
                        sheetOrder);
                assertEquals(vector.path("expected").asText(), actual, vector.path("id").asText());
            }
        }
    }
}
