package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

import java.time.Instant;

@Entity
@Table(name = "user_preference")
public class UserPreferenceEntity {
    @Id
    @Column(name = "subject", nullable = false, length = 500)
    private String subject;
    @Column(name = "default_space_id", length = 200)
    private String defaultSpaceId;
    @Column(name = "default_folder_id", length = 200)
    private String defaultFolderId;
    @Column(name = "auto_save", nullable = false)
    private boolean autoSave;
    @Column(name = "auto_sync", nullable = false)
    private boolean autoSync;
    @Column(name = "offline_cache", nullable = false)
    private boolean offlineCache;
    @Column(name = "import_compatibility", nullable = false, length = 32)
    private String importCompatibility;
    @Column(name = "language", length = 32)
    private String language;
    @Column(name = "theme", length = 32)
    private String theme;
    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected UserPreferenceEntity() {}

    public UserPreferenceEntity(String subject, String defaultSpaceId, String defaultFolderId, boolean autoSave,
                                boolean autoSync, boolean offlineCache, String importCompatibility, String language,
                                String theme, Instant updatedAt) {
        this.subject = subject;
        this.defaultSpaceId = defaultSpaceId;
        this.defaultFolderId = defaultFolderId;
        this.autoSave = autoSave;
        this.autoSync = autoSync;
        this.offlineCache = offlineCache;
        this.importCompatibility = importCompatibility;
        this.language = language;
        this.theme = theme;
        this.updatedAt = updatedAt;
    }

    public String getSubject() { return subject; }
    public String getDefaultSpaceId() { return defaultSpaceId; }
    public String getDefaultFolderId() { return defaultFolderId; }
    public boolean isAutoSave() { return autoSave; }
    public boolean isAutoSync() { return autoSync; }
    public boolean isOfflineCache() { return offlineCache; }
    public String getImportCompatibility() { return importCompatibility; }
    public String getLanguage() { return language; }
    public String getTheme() { return theme; }
    public Instant getUpdatedAt() { return updatedAt; }
}
