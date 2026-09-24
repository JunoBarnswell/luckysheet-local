package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonProperty;

/** Immutable descriptor for a server-materialized query data-source block. */
public record QueryDataSourceBlock(
        @JsonProperty("blockId") String blockId,
        @JsonProperty("startRow") int startRow,
        @JsonProperty("rowCount") int rowCount,
        @JsonProperty("checksum") String checksum,
        @JsonProperty("byteLength") int byteLength,
        @JsonProperty("encoding") String encoding
) {
    public QueryDataSourceBlock {
        if (blockId == null || blockId.isBlank() || blockId.length() > 200 || startRow < 0 || rowCount < 1
                || (long) startRow + rowCount > Integer.MAX_VALUE
                || checksum == null || !checksum.matches("[A-Fa-f0-9]{64}")
                || byteLength < 1 || encoding == null || !encoding.equals("columnar-v1")) {
            throw new IllegalArgumentException("Query data-source block metadata is invalid");
        }
    }
}
