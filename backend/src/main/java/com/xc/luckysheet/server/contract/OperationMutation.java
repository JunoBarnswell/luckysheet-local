package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

public record OperationMutation(
        @JsonProperty("id") String id,
        @JsonProperty("sheetId") String sheetId,
        @JsonProperty("params") JsonNode params
) {
    @JsonCreator
    public OperationMutation {
        if (id == null || id.isBlank()) throw new IllegalArgumentException("mutation id is required");
        if (sheetId == null || sheetId.isBlank()) throw new IllegalArgumentException("mutation sheetId is required");
        if (params == null) throw new IllegalArgumentException("mutation params are required");
    }
}
