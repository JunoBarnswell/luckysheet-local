package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record FolderResponse(
        String folderId,
        String spaceId,
        String parentFolderId,
        String name,
        Instant createdAt,
        Instant updatedAt
) {
}
