package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record RangeRef(
        @JsonProperty("sheetId") String sheetId,
        @JsonProperty("startRow") int startRow,
        @JsonProperty("endRow") int endRow,
        @JsonProperty("startColumn") int startColumn,
        @JsonProperty("endColumn") int endColumn
) {
    @JsonCreator
    public RangeRef {
        if (sheetId == null || sheetId.isBlank()) throw new IllegalArgumentException("sheetId is required");
        if (startRow < 0 || endRow < startRow || startColumn < 0 || endColumn < startColumn) {
            throw new IllegalArgumentException("Invalid range bounds");
        }
    }
}
