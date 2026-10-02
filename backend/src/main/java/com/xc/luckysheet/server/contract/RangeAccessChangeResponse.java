package com.xc.luckysheet.server.contract;

public record RangeAccessChangeResponse(RangeAccessRegion region, long accessRevision) {
    public RangeAccessChangeResponse {
        if (region == null || accessRevision < 1) throw new IllegalArgumentException("Range access change is invalid");
    }
}
