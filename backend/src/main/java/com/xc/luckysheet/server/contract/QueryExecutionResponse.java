package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

import java.time.Instant;
import java.util.List;

public record QueryExecutionResponse(
        @JsonProperty("queryId") String queryId,
        @JsonProperty("connectorId") String connectorId,
        @JsonProperty("sourceRef") String sourceRef,
        @JsonProperty("sourceRevision") long sourceRevision,
        @JsonProperty("columns") List<String> columns,
        @JsonProperty("rows") List<List<JsonNode>> rows,
        @JsonProperty("rowCount") long rowCount,
        @JsonProperty("executedAt") Instant executedAt,
        @JsonProperty("durationMs") long durationMs
) {
    public QueryExecutionResponse {
        columns = List.copyOf(columns);
        rows = rows.stream().map(row -> row.stream().<JsonNode>map(value -> value == null ? null : value.deepCopy()).toList()).toList();
    }
}
