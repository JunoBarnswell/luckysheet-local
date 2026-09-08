package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

/** Subsequent analytics page requests must carry the server-issued token. */
public record AnalyticsExecutionRequest(
        @JsonProperty("executionToken") String executionToken,
        @JsonProperty("request") JsonNode request
) {
    @JsonCreator
    public AnalyticsExecutionRequest {
        if (executionToken == null || !executionToken.matches("[0-9a-fA-F-]{36}")) {
            throw new IllegalArgumentException("executionToken is invalid");
        }
        if (request == null || !request.isObject()) throw new IllegalArgumentException("analytics request must be an object");
        if (!request.path("kind").isTextual() || !java.util.Set.of("filter", "query", "pivot").contains(request.path("kind").asText())) {
            throw new IllegalArgumentException("analytics request kind must be filter, query or pivot");
        }
        if (!request.path("revision").isIntegralNumber() || request.path("revision").asLong(-1) < 0) {
            throw new IllegalArgumentException("analytics request revision is required");
        }
        request = request.deepCopy();
    }
}
