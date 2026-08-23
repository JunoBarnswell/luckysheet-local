package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record CreateSpaceRequest(
        @JsonProperty("name") String name,
        @JsonProperty("kind") WorkspaceSpaceType kind
) {
    @JsonCreator
    public CreateSpaceRequest {
        if (name == null || name.isBlank() || name.length() > 200) throw new IllegalArgumentException("Space name is required");
        if (kind == null) throw new IllegalArgumentException("Space kind is required");
    }
}
