package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;

import java.util.List;

public interface MutationDescriptor {
    String id();

    boolean internalOnly();

    List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation);

    JsonNode apply(JsonNode snapshot, OperationMutation mutation);
}
