package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

public record CreateWorkbookRequest(
        @JsonProperty("unitId") String unitId,
        @JsonProperty("name") String name,
        @JsonProperty("snapshot") JsonNode snapshot
) {
    @JsonCreator
    public CreateWorkbookRequest {
        if (unitId == null || unitId.isBlank() || name == null || name.isBlank() || snapshot == null || !snapshot.isObject()) {
            throw new IllegalArgumentException("unitId, name and object snapshot are required");
        }
    }
}
