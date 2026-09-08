package com.xc.luckysheet.server.contract;

import java.time.Instant;

/** Actor-global preferences. Workbook-specific state is intentionally separate. */
public record UserPreferences(
        String defaultSpaceId,
        String defaultFolderId,
        boolean autoSave,
        boolean autoSync,
        String importCompatibility,
        String language,
        String theme,
        Instant updatedAt
) {
}
