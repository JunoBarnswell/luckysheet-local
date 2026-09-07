package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import java.time.Instant;

public record WorkbookArtifactResponse(
        String unitId,
        String fileName,
        String mimeType,
        String checksum,
        long revision,
        long byteLength,
        JsonNode nativeMetadata,
        Instant createdAt,
        Instant updatedAt
) {
}
