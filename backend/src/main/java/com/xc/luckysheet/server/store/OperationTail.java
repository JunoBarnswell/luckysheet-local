package com.xc.luckysheet.server.store;

import java.util.List;

/**
 * A bounded, explicitly addressed operation interval.
 *
 * <p>The workbook store never exposes an unbounded operation-log read.  The
 * interval is part of the value so a caller cannot accidentally replay rows
 * that belong to another checkpoint window.</p>
 */
public record OperationTail(
        String unitId,
        long fromExclusive,
        long toInclusive,
        List<OperationRow> operations,
        long envelopeBytes
) {
    public OperationTail {
        if (unitId == null || unitId.isBlank()) throw new IllegalArgumentException("Operation tail unitId is required");
        if (fromExclusive < 0 || toInclusive < fromExclusive) throw new IllegalArgumentException("Operation tail interval is invalid");
        operations = List.copyOf(operations);
        if (envelopeBytes < 0) throw new IllegalArgumentException("Operation tail byte count is invalid");
    }

    public static OperationTail empty(String unitId, long fromExclusive, long toInclusive) {
        return new OperationTail(unitId, fromExclusive, toInclusive, List.of(), 0);
    }
}
