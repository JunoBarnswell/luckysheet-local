package com.xc.luckysheet.server.store;

import java.time.Instant;
import java.util.UUID;

public record OutboxRow(
        UUID eventId,
        String unitId,
        String operationId,
        long revision,
        String payloadJson,
        Instant createdAt,
        int attempts
) {
}
