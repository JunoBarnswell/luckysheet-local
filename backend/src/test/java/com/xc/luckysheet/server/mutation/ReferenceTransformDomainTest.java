package com.xc.luckysheet.server.mutation;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class ReferenceTransformDomainTest {
    @Test
    void mapsPointsAndInclusiveIntervalsWithSharedInsertDeleteSemantics() {
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.MAPPED, 2),
                ReferenceTransformDomain.mapPoint(2, 3, 2, true, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.MAPPED, 5),
                ReferenceTransformDomain.mapPoint(3, 3, 2, true, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.DELETED, -1),
                ReferenceTransformDomain.mapPoint(3, 3, 2, false, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.PointMapping(ReferenceTransformDomain.PointKind.MAPPED, 3),
                ReferenceTransformDomain.mapPoint(5, 3, 2, false, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.IntervalMapping(ReferenceTransformDomain.IntervalKind.MAPPED, 2, 6),
                ReferenceTransformDomain.mapInterval(2, 4, 3, 2, true, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.IntervalMapping(ReferenceTransformDomain.IntervalKind.MAPPED, 2, 3),
                ReferenceTransformDomain.mapInterval(2, 5, 3, 2, false, ReferenceTransformDomain.MAX_ROW_INDEX));
        assertEquals(new ReferenceTransformDomain.IntervalMapping(ReferenceTransformDomain.IntervalKind.DELETED, -1, -1),
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
}
