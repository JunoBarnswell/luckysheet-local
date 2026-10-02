package com.xc.luckysheet.server.contract;

import java.util.List;

public record RangeAccessRegionRequest(
        String sheetId,
        RangeRef range,
        RangeAccessLevel defaultAccess,
        List<RangeAccessGrant> grants
) {
    public RangeAccessRegionRequest {
        grants = grants == null ? List.of() : List.copyOf(grants);
    }
}
