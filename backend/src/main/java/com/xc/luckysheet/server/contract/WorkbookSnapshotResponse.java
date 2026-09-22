package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

public record WorkbookSnapshotResponse(
        @JsonProperty("unitId") String unitId,
        @JsonProperty("snapshot") JsonNode snapshot,
        @JsonProperty("revision") long revision,
        @JsonProperty("checksum") String checksum
) {
    @JsonCreator
    public WorkbookSnapshotResponse {
        if (unitId == null || unitId.isBlank() || snapshot == null || revision < 0 || checksum == null || checksum.isBlank()) {
            throw new IllegalArgumentException("Invalid snapshot response");
        }
    }
}
