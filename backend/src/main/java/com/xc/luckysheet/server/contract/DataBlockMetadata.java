package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record DataBlockMetadata(
        String unitId,
        String sourceId,
        String blockId,
        String checksum,
        int byteLength,
        Instant updatedAt
) {
}
