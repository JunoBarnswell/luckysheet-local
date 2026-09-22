package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record WorkbookUserState(
        String unitId,
        boolean favorite,
        Instant lastOpenedAt,
        boolean autoSave,
        boolean autoSync,
        String defaultCreateLocation,
        String importCompatibilityLevel,
        String language,
        boolean offlineCache,
        String theme,
        Instant updatedAt
) {
    public WorkbookUserState(String unitId, boolean favorite, Instant lastOpenedAt, Instant updatedAt) {
        this(unitId, favorite, lastOpenedAt, true, true, "remote", "standard", null, true, "system", updatedAt);
    }
}
