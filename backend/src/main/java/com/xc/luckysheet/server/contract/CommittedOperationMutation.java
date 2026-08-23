package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;
import java.util.List;

public record CommittedOperationMutation(
        @JsonProperty("id") String id,
        @JsonProperty("sheetId") String sheetId,
        @JsonProperty("params") JsonNode params,
        @JsonProperty("affectedRanges") List<RangeRef> affectedRanges
) {
    @JsonCreator
    public CommittedOperationMutation {
        if (id == null || id.isBlank()) throw new IllegalArgumentException("mutation id is required");
        if (sheetId == null || sheetId.isBlank()) throw new IllegalArgumentException("mutation sheetId is required");
        if (params == null) throw new IllegalArgumentException("mutation params are required");
        if (affectedRanges == null) throw new IllegalArgumentException("affectedRanges are server-owned");
        affectedRanges = List.copyOf(affectedRanges);
    }

    public static CommittedOperationMutation from(OperationMutation mutation, List<RangeRef> ranges) {
        return new CommittedOperationMutation(mutation.id(), mutation.sheetId(), mutation.params(), ranges);
    }
}
