package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record AssetMetadata(
        String schema,
        String unitId,
        String assetId,
        String contentHash,
        String mimeType,
        int byteLength,
        Integer width,
        Integer height,
        Instant updatedAt
) {
}
