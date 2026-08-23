package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record CopyWorkbookRequest(
        @JsonProperty("name") String name,
        @JsonProperty("spaceId") String spaceId,
        @JsonProperty("folderId") String folderId
) {
    @JsonCreator
    public CopyWorkbookRequest {
        if (name != null && name.length() > 500) throw new IllegalArgumentException("Workbook name is too long");
    }
}
