package com.xc.luckysheet.server.contract;

import java.time.Instant;
import java.util.List;

public record RangeAccessRegion(
        String id,
        String unitId,
        String sheetId,
        RangeRef range,
        RangeAccessLevel defaultAccess,
        List<RangeAccessGrant> grants,
        String createdBy,
        Instant createdAt,
        Instant updatedAt
) {
    public RangeAccessRegion {
        grants = grants == null ? List.of() : List.copyOf(grants);
    }
}
