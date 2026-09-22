package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonProperty;

import java.time.Instant;
import java.util.List;

/** Metadata for an explicit server-query block session. */
public record QueryBlockExecutionResponse(
        @JsonProperty("queryId") String queryId,
        @JsonProperty("executionId") String executionId,
        @JsonProperty("connectorId") String connectorId,
        @JsonProperty("sourceRef") String sourceRef,
        @JsonProperty("sourceRevision") long sourceRevision,
        @JsonProperty("columns") List<String> columns,
        @JsonProperty("columnTypes") List<String> columnTypes,
        @JsonProperty("rowCount") long rowCount,
        @JsonProperty("blockRowCount") int blockRowCount,
        @JsonProperty("executedAt") Instant executedAt,
        @JsonProperty("durationMs") long durationMs
) {
    public QueryBlockExecutionResponse {
        columns = List.copyOf(columns);
        columnTypes = List.copyOf(columnTypes);
        if (queryId == null || queryId.isBlank() || executionId == null || executionId.isBlank()
                || connectorId == null || connectorId.isBlank() || sourceRef == null || sourceRef.isBlank()) {
            throw new IllegalArgumentException("Query block execution identity is invalid");
        }
        if (columns.isEmpty() || columns.size() != columnTypes.size() || rowCount < 0 || blockRowCount < 1
                || sourceRevision < 0 || executedAt == null || durationMs < 0) {
            throw new IllegalArgumentException("Query block execution metadata is inconsistent");
        }
    }
}
