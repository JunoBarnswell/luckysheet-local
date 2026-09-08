package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

/** A revision-pinned native analytics task declaration. */
public record AnalyticsPrepareRequest(
        @JsonProperty("queryId") String queryId,
        @JsonProperty("revision") long revision,
        @JsonProperty("request") JsonNode request
) {
    @JsonCreator
    public AnalyticsPrepareRequest {
        if (queryId == null || !queryId.matches("[A-Za-z0-9._:-]{1,200}")) {
            throw new IllegalArgumentException("queryId is invalid");
        }
        if (revision < 0) throw new IllegalArgumentException("revision must be non-negative");
        if (request == null || !request.isObject()) throw new IllegalArgumentException("analytics request must be an object");
        if (!request.path("kind").isTextual() || !java.util.Set.of("filter", "query", "pivot").contains(request.path("kind").asText())) {
            throw new IllegalArgumentException("analytics request kind must be filter, query or pivot");
        }
        if (!request.path("revision").isIntegralNumber() || request.path("revision").asLong(-1) != revision) {
            throw new IllegalArgumentException("analytics request revision must equal the pinned revision");
        }
        request = request.deepCopy();
    }
}
