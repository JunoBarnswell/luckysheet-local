package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;
import java.time.Instant;

/** A bounded native analytics page plus its server-owned proof identity. */
public record AnalyticsExecutionResponse(
        @JsonProperty("queryId") String queryId,
        @JsonProperty("sourceRevision") long sourceRevision,
        @JsonProperty("executionToken") String executionToken,
        @JsonProperty("resultHash") String resultHash,
        @JsonProperty("result") JsonNode result,
        @JsonProperty("executedAt") Instant executedAt,
        @JsonProperty("durationMs") long durationMs
) {
    public AnalyticsExecutionResponse {
        if (result == null || !result.isObject()) throw new IllegalArgumentException("analytics result must be an object");
        result = result.deepCopy();
    }
}
