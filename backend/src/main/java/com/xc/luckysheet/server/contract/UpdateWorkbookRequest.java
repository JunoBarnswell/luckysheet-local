package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record UpdateWorkbookRequest(
        @JsonProperty("name") String name,
        @JsonProperty("spaceId") String spaceId,
        @JsonProperty("folderId") String folderId
) {
    @JsonCreator
    public UpdateWorkbookRequest {
        if ((name == null || name.isBlank()) && (spaceId == null || spaceId.isBlank()) && (folderId == null || folderId.isBlank())) {
            throw new IllegalArgumentException("At least one workbook metadata field is required");
        }
        if (name != null && name.length() > 500) throw new IllegalArgumentException("Workbook name is too long");
    }
}
