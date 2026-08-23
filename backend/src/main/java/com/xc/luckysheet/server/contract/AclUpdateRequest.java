package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record AclUpdateRequest(@JsonProperty("role") WorkbookAclRole role) {
    @JsonCreator
    public AclUpdateRequest {
        if (role == null || role == WorkbookAclRole.OWNER) throw new IllegalArgumentException("Only editor, commenter or viewer may be granted");
    }
}
