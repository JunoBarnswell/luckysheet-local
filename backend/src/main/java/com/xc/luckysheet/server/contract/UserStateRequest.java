package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.time.Instant;

public record UserStateRequest(
        @JsonProperty("favorite") Boolean favorite,
        @JsonProperty("lastOpenedAt") Instant lastOpenedAt,
        @JsonProperty("autoSave") Boolean autoSave,
        @JsonProperty("autoSync") Boolean autoSync,
        @JsonProperty("defaultCreateLocation") String defaultCreateLocation,
        @JsonProperty("importCompatibilityLevel") String importCompatibilityLevel,
        @JsonProperty("language") String language,
        @JsonProperty("offlineCache") Boolean offlineCache,
        @JsonProperty("theme") String theme
) {
    public UserStateRequest(Boolean favorite, Instant lastOpenedAt) {
        this(favorite, lastOpenedAt, null, null, null, null, null, null, null);
    }

    @JsonCreator
    public UserStateRequest {
        if (favorite == null && lastOpenedAt == null && autoSave == null && autoSync == null
                && defaultCreateLocation == null && importCompatibilityLevel == null && language == null
                && offlineCache == null && theme == null) throw new IllegalArgumentException("User state is empty");
        if (defaultCreateLocation != null && !java.util.Set.of("local", "remote").contains(defaultCreateLocation)) {
            throw new IllegalArgumentException("defaultCreateLocation must be local or remote");
        }
        if (importCompatibilityLevel != null && !java.util.Set.of("standard", "strict").contains(importCompatibilityLevel)) {
            throw new IllegalArgumentException("importCompatibilityLevel is invalid");
        }
        if (theme != null && !java.util.Set.of("light", "system").contains(theme)) throw new IllegalArgumentException("theme is invalid");
    }
}
