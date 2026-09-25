package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

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
            assertEquals(1, vectors.path("version").asInt());

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
        }
    }

    private static int maximum(JsonNode vector) {
        return switch (vector.path("axis").asText()) {
            case "row" -> ReferenceTransformDomain.MAX_ROW_INDEX;
            case "column" -> ReferenceTransformDomain.MAX_COLUMN_INDEX;
            default -> throw new IllegalArgumentException("Shared reference vector axis is invalid");
        };
    }
}
