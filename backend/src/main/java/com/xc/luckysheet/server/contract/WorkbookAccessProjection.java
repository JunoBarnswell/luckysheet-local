package com.xc.luckysheet.server.contract;

/** Server-computed role and operation cursor for the current authenticated subject. */
public record WorkbookAccessProjection(String unitId, WorkbookAclRole role, long nextClientSequence) {
    public WorkbookAccessProjection {
        if (unitId == null || unitId.isBlank() || role == null || nextClientSequence < 1) {
            throw new IllegalArgumentException("Workbook access projection is invalid");
        }
    }
}
