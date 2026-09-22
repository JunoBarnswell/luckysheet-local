package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

import java.util.List;

/** One bounded page from an explicit server-query block session. */
public record QueryBlockResponse(
        @JsonProperty("queryId") String queryId,
        @JsonProperty("executionId") String executionId,
        @JsonProperty("offset") long offset,
        @JsonProperty("rows") List<List<JsonNode>> rows,
        @JsonProperty("hasMore") boolean hasMore
) {
    public QueryBlockResponse {
        if (queryId == null || queryId.isBlank() || executionId == null || executionId.isBlank() || offset < 0 || rows == null) {
            throw new IllegalArgumentException("Query block response metadata is invalid");
        }
        // QueryTable owns these nodes for the short lifetime of the session;
        // the HTTP serializer is read-only. Copying every scalar here doubles
        // heap and CPU for each large block response without adding isolation.
        rows = rows.stream().map(row -> java.util.Collections.unmodifiableList(new java.util.ArrayList<>(row))).toList();
    }
}
