package com.xc.luckysheet.server.mutation;

/** Zero-based point mapping shared by server cell and reference transforms. */
final class StructuralAxisCoordinate {
    private StructuralAxisCoordinate() {
    }

    static long mapPoint(int position, int at, int count, boolean insert) {
        if (insert) return position < at ? position : (long) position + count;
        long deletedEnd = (long) at + count - 1;
        if (position < at) return position;
        if (position > deletedEnd) return (long) position - count;
        return -1;
    }
}
