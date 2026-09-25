package com.xc.luckysheet.server.mutation;

/** Canonical server-side point and inclusive-interval mapping for worksheet references. */
final class ReferenceTransformDomain {
    static final int MAX_ROW_INDEX = 1_048_575;
    static final int MAX_COLUMN_INDEX = 16_383;

    enum PointKind { MAPPED, DELETED, OUT_OF_BOUNDS }

    record PointMapping(PointKind kind, Long position) {
    }

    enum IntervalKind { MAPPED, DELETED, OUT_OF_BOUNDS }

    record IntervalMapping(IntervalKind kind, Long start, Long end) {
    }

    private ReferenceTransformDomain() {
    }

    static PointMapping mapPoint(int position, int at, int count, boolean insert, int maximum) {
        if (position < 0 || at < 0 || count < 1 || maximum < 0 || at > maximum || (long) count > (long) maximum + 1) {
            throw new IllegalArgumentException("Reference transform point inputs are invalid");
        }
        if (position > maximum) return new PointMapping(PointKind.OUT_OF_BOUNDS, (long) position);
        if (!insert && position >= at && (long) position - at < count) {
            return new PointMapping(PointKind.DELETED, null);
        }
        long mapped = insert
                ? position < at ? position : (long) position + count
                : position < at ? position : (long) position - count;
        return mapped <= maximum
                ? new PointMapping(PointKind.MAPPED, mapped)
                : new PointMapping(PointKind.OUT_OF_BOUNDS, mapped);
    }

    static IntervalMapping mapInterval(int start, int end, int at, int count, boolean insert, int maximum) {
        if (start < 0 || end < 0 || at < 0 || count < 1 || maximum < 0
                || start > maximum || end > maximum || at > maximum || (long) count > (long) maximum + 1
                || (!insert && (long) count - 1 > (long) maximum - at)) {
            throw new IllegalArgumentException("Reference transform interval inputs are invalid");
        }
        long low = Math.min(start, end);
        long high = Math.max(start, end);
        long nextStart;
        long nextEnd;
        if (insert) {
            if (at <= low) {
                nextStart = low + count;
                nextEnd = high + count;
            } else if (at <= high) {
                nextStart = low;
                nextEnd = high + count;
            } else {
                nextStart = low;
                nextEnd = high;
            }
        } else {
            long deletedEnd = (long) at + count - 1;
            if (high < at) {
                nextStart = low;
                nextEnd = high;
            } else if (low > deletedEnd) {
                nextStart = low - count;
                nextEnd = high - count;
            } else {
                nextStart = low < at ? low : at;
                nextEnd = high > deletedEnd ? high - count : (long) at - 1;
            }
        }
        if (nextStart > nextEnd) return new IntervalMapping(IntervalKind.DELETED, null, null);
        if (nextStart < 0 || nextEnd > maximum) {
            return new IntervalMapping(IntervalKind.OUT_OF_BOUNDS, nextStart, nextEnd);
        }
        return new IntervalMapping(IntervalKind.MAPPED, nextStart, nextEnd);
    }
}
