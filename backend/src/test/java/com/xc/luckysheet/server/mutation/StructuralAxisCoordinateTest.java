package com.xc.luckysheet.server.mutation;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

class StructuralAxisCoordinateTest {
    @Test
    void insertionLeavesEarlierPointsAndMovesTheInsertionPointAndFollowingPoints() {
        assertEquals(2, StructuralAxisCoordinate.mapPoint(2, 3, 2, true));
        assertEquals(5, StructuralAxisCoordinate.mapPoint(3, 3, 2, true));
        assertEquals(10, StructuralAxisCoordinate.mapPoint(8, 3, 2, true));
    }

    @Test
    void deletionPreservesEarlierPointsRemovesItsIntervalAndClosesTheFollowingGap() {
        assertEquals(2, StructuralAxisCoordinate.mapPoint(2, 3, 2, false));
        assertEquals(-1, StructuralAxisCoordinate.mapPoint(3, 3, 2, false));
        assertEquals(-1, StructuralAxisCoordinate.mapPoint(4, 3, 2, false));
        assertEquals(3, StructuralAxisCoordinate.mapPoint(5, 3, 2, false));
    }
}
