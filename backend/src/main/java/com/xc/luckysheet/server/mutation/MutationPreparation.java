package com.xc.luckysheet.server.mutation;

import com.xc.luckysheet.server.contract.RangeRef;

import java.util.List;

/** A single server-resolved mutation, ready to reduce into a snapshot. */
public record MutationPreparation(MutationDescriptor descriptor, List<RangeRef> affectedRanges) {
    public MutationPreparation {
        affectedRanges = List.copyOf(affectedRanges);
    }
}
