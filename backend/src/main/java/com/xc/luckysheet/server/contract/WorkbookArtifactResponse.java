package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record WorkbookArtifactResponse(
        String unitId,
        String fileName,
        String mimeType,
        String checksum,
        long byteLength,
        Instant createdAt,
        Instant updatedAt
) {
}
