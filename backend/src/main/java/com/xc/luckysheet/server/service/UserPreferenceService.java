package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.UserPreferences;
import com.xc.luckysheet.server.contract.UserPreferencesRequest;
import com.xc.luckysheet.server.persistence.UserPreferenceEntity;
import com.xc.luckysheet.server.persistence.UserPreferenceEntityRepository;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.Instant;

@Service
public class UserPreferenceService {
    private final UserPreferenceEntityRepository preferences;

    public UserPreferenceService(UserPreferenceEntityRepository preferences) {
        this.preferences = preferences;
    }

    public UserPreferences get(String subject) {
        if (subject == null || subject.isBlank()) throw ServiceException.unauthenticated("A registered actor is required");
        return preferences.findById(subject).map(this::response)
                .orElseGet(() -> new UserPreferences(null, null, true, true, true, "B", null, "system", null));
    }

    @Transactional
    public UserPreferences update(String subject, UserPreferencesRequest request) {
        if (subject == null || subject.isBlank()) throw ServiceException.unauthenticated("A registered actor is required");
        Instant now = Instant.now();
        UserPreferenceEntity entity = preferences.findById(subject).orElseGet(() ->
                new UserPreferenceEntity(subject, null, null, true, true, true, "B", null, "system", now));
        entity.update(normalizeId(request.defaultSpaceId()), normalizeId(request.defaultFolderId()), request.autoSave(), request.autoSync(),
                request.offlineCache(), normalizeCompatibility(request.importCompatibility()), normalizeLanguage(request.language()),
                normalizeTheme(request.theme()), now);
        preferences.save(entity);
        return response(entity);
    }

    private String normalizeId(String value) {
        if (value == null) return null;
        String normalized = value.trim();
        return normalized.isEmpty() ? null : normalized;
    }

    private String normalizeLanguage(String value) {
        if (value == null) return null;
        String normalized = value.trim();
        return normalized.isEmpty() ? null : normalized;
    }

    private String normalizeCompatibility(String value) {
        if (value == null) return null;
        return value.trim().toUpperCase(java.util.Locale.ROOT);
    }

    private String normalizeTheme(String value) {
        if (value == null) return null;
        return value.trim().toLowerCase(java.util.Locale.ROOT);
    }

    private UserPreferences response(UserPreferenceEntity entity) {
        return new UserPreferences(entity.getDefaultSpaceId(), entity.getDefaultFolderId(), entity.isAutoSave(), entity.isAutoSync(),
                entity.isOfflineCache(), entity.getImportCompatibility(), entity.getLanguage(), entity.getTheme(), entity.getUpdatedAt());
    }
}
