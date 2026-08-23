package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record SpaceMemberRequest(@JsonProperty("role") WorkbookAclRole role) {
    @JsonCreator
    public SpaceMemberRequest {
        if (role == null || role == WorkbookAclRole.OWNER) throw new IllegalArgumentException("Member role must be editor, commenter or viewer");
    }
}
