package com.xc.luckysheet.server.store;

import java.time.Instant;

public record DataBlockRow(
        String unitId,
        String sourceId,
        String blockId,
        String checksum,
        int byteLength,
        byte[] content,
        Instant createdAt,
        Instant updatedAt
) {
}
