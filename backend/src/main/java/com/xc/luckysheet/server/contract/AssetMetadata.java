package com.xc.luckysheet.server.contract;

import java.time.Instant;
import com.fasterxml.jackson.annotation.JsonInclude;

@JsonInclude(JsonInclude.Include.NON_NULL)
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
