package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record FolderRequest(
        @JsonProperty("name") String name,
        @JsonProperty("parentFolderId") String parentFolderId
) {
    @JsonCreator
    public FolderRequest {
        if (name == null || name.isBlank() || name.length() > 200) throw new IllegalArgumentException("Folder name is required");
    }
}
