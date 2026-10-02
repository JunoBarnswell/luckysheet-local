package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record SpaceMemberResponse(
        String spaceId,
        String subject,
        WorkbookRole role,
        Instant createdAt,
        Instant updatedAt
) {
}
