package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record SpaceMemberResponse(
        String spaceId,
        String subject,
        WorkbookAclRole role,
        Instant createdAt,
        Instant updatedAt
) {
}
