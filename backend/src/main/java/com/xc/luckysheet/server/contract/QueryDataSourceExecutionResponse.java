package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonProperty;

import java.time.Instant;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

/** Metadata for a query whose bounded blocks are already durable data blocks. */
public record QueryDataSourceExecutionResponse(
        @JsonProperty("queryId") String queryId,
        @JsonProperty("connectorId") String connectorId,
        @JsonProperty("sourceRef") String sourceRef,
        @JsonProperty("sourceRevision") long sourceRevision,
        @JsonProperty("columns") List<String> columns,
        @JsonProperty("columnTypes") List<String> columnTypes,
        @JsonProperty("rowCount") long rowCount,
        @JsonProperty("blockRowCount") int blockRowCount,
        @JsonProperty("blocks") List<QueryDataSourceBlock> blocks,
        @JsonProperty("executedAt") Instant executedAt,
        @JsonProperty("durationMs") long durationMs
) {
    private static final int MAX_BLOCK_COUNT = 10_000;

    public QueryDataSourceExecutionResponse {
        columns = List.copyOf(columns);
        columnTypes = List.copyOf(columnTypes);
        blocks = List.copyOf(blocks);
        if (queryId == null || queryId.isBlank() || connectorId == null || connectorId.isBlank()
                || sourceRef == null || sourceRef.isBlank() || sourceRevision < 0
                || columns.isEmpty() || columns.size() > 16_384 || columns.size() != columnTypes.size() || rowCount < 0
                || blocks.size() > MAX_BLOCK_COUNT
                || blockRowCount < 1 || executedAt == null || durationMs < 0) {
            throw new IllegalArgumentException("Query data-source execution metadata is inconsistent");
        }
        Set<String> columnNames = new HashSet<>();
        for (int index = 0; index < columns.size(); index++) {
            String column = columns.get(index);
            String type = columnTypes.get(index);
            if (column == null || column.isBlank() || column.length() > 200 || !columnNames.add(column)
                    || type == null || !Set.of("text", "number", "boolean", "date", "mixed").contains(type)) {
                throw new IllegalArgumentException("Query data-source field schema is invalid");
            }
        }
        Set<String> blockIds = new HashSet<>();
        long covered = 0;
        for (QueryDataSourceBlock block : blocks) {
            if (!blockIds.add(block.blockId()) || block.startRow() != covered || block.rowCount() > blockRowCount) {
                throw new IllegalArgumentException("Query data-source blocks do not provide contiguous coverage");
            }
            covered += block.rowCount();
        }
        if (covered != rowCount) throw new IllegalArgumentException("Query data-source blocks do not cover the result");
    }
}
