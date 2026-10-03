package com.xc.luckysheet.server.store;

import com.xc.luckysheet.server.contract.WorkbookRole;

import java.time.Instant;
import java.util.UUID;

public record ShareRow(
        UUID shareId,
        String unitId,
        String tokenHash,
        WorkbookRole role,
        Instant expiresAt,
        Instant revokedAt,
        String createdBy,
        Instant createdAt
) {
}
