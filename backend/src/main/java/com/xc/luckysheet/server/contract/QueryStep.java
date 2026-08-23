package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record QueryStep(
        @JsonProperty("id") String id,
        @JsonProperty("kind") String kind,
        @JsonProperty("name") String name,
        @JsonProperty("config") JsonNode config,
        @JsonProperty("enabled") boolean enabled
) {
    @JsonCreator
    public QueryStep {
        if (id == null || id.isBlank()) throw new IllegalArgumentException("Query step id is required");
        if (kind == null || kind.isBlank()) throw new IllegalArgumentException("Query step kind is required");
        if (name == null || name.isBlank()) throw new IllegalArgumentException("Query step name is required");
        if (config == null || !config.isObject()) throw new IllegalArgumentException("Query step config must be an object");
        config = config.deepCopy();
    }
}
