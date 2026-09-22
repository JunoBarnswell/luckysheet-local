package com.xc.luckysheet.server.contract;

/** Server-computed role projection for the current authenticated subject. */
public record WorkbookAccessProjection(String unitId, WorkbookAclRole role) {
    public WorkbookAccessProjection {
        if (unitId == null || unitId.isBlank() || role == null) {
            throw new IllegalArgumentException("Workbook access projection is invalid");
        }
    }
}
