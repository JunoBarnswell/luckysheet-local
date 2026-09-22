package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record UserPreferencesRequest(
        @JsonProperty("defaultSpaceId") String defaultSpaceId,
        @JsonProperty("defaultFolderId") String defaultFolderId,
        @JsonProperty("autoSave") Boolean autoSave,
        @JsonProperty("autoSync") Boolean autoSync,
        @JsonProperty("offlineCache") Boolean offlineCache,
        @JsonProperty("importCompatibility") String importCompatibility,
        @JsonProperty("language") String language,
        @JsonProperty("theme") String theme
) {
    @JsonCreator
    public UserPreferencesRequest {
        if (defaultSpaceId == null && defaultFolderId == null && autoSave == null && autoSync == null && offlineCache == null
                && importCompatibility == null && language == null && theme == null) {
            throw new IllegalArgumentException("User preferences are empty");
        }
        if (defaultSpaceId != null && defaultSpaceId.length() > 200) {
            throw new IllegalArgumentException("defaultSpaceId is too long");
        }
        if (defaultFolderId != null && defaultFolderId.length() > 200) {
            throw new IllegalArgumentException("defaultFolderId is too long");
        }
        if (language != null && language.length() > 32) {
            throw new IllegalArgumentException("language is too long");
        }
        if (importCompatibility != null && !java.util.Set.of("A", "B", "C").contains(importCompatibility.trim().toUpperCase(java.util.Locale.ROOT))) {
            throw new IllegalArgumentException("importCompatibility is invalid");
        }
        if (theme != null && !java.util.Set.of("light", "dark", "system").contains(theme.trim().toLowerCase(java.util.Locale.ROOT))) {
            throw new IllegalArgumentException("theme is invalid");
        }
    }
}
