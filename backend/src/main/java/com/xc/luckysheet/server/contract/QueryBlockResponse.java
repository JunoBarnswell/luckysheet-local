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
        rows = rows.stream().map(row -> row.stream().<JsonNode>map(value -> value == null ? null : value.deepCopy()).toList()).toList();
    }
}
