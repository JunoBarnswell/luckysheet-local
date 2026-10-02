package com.xc.luckysheet.server.contract;

import java.util.List;

/** Server-computed role projection for the current authenticated subject. */
public record WorkbookAccessProjection(
        String unitId,
        WorkbookRole role,
        long accessRevision,
        List<EffectiveAccessRegion> regions
) {
    public WorkbookAccessProjection {
        if (unitId == null || unitId.isBlank() || role == null) {
            throw new IllegalArgumentException("Workbook access projection is invalid");
        }
        if (accessRevision < 0) throw new IllegalArgumentException("Access revision is invalid");
        regions = regions == null ? List.of() : List.copyOf(regions);
    }
}
