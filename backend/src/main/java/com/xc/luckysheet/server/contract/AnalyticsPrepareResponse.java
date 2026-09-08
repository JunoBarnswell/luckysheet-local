package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonProperty;
import java.time.Instant;

public record AnalyticsPrepareResponse(
        @JsonProperty("queryId") String queryId,
        @JsonProperty("sourceRevision") long sourceRevision,
        @JsonProperty("executionToken") String executionToken,
        @JsonProperty("expiresAt") Instant expiresAt
) { }
