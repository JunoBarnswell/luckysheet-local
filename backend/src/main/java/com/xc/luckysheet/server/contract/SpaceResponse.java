package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record SpaceResponse(
        String spaceId,
        String name,
        WorkspaceSpaceType kind,
        String createdBy,
        WorkbookRole role,
        Instant createdAt,
        Instant updatedAt
) {
}
