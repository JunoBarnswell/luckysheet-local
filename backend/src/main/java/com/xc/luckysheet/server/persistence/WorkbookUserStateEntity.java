package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Embeddable;
import jakarta.persistence.EmbeddedId;
import jakarta.persistence.Entity;
import jakarta.persistence.Index;
import jakarta.persistence.Table;

import java.io.Serializable;
import java.time.Instant;
import java.util.Objects;

@Entity
@Table(name = "workbook_user_state", indexes = @Index(name = "workbook_user_state_subject_idx", columnList = "subject,unit_id"))
public class WorkbookUserStateEntity {
    @EmbeddedId
    private Id id;

    @Column(name = "favorite", nullable = false)
    private boolean favorite;

    @Column(name = "last_opened_at")
    private Instant lastOpenedAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    @Column(name = "auto_save", nullable = false)
    private boolean autoSave;
    @Column(name = "auto_sync", nullable = false)
    private boolean autoSync;
    @Column(name = "default_create_location", nullable = false, length = 16)
    private String defaultCreateLocation;
    @Column(name = "import_compatibility_level", nullable = false, length = 16)
    private String importCompatibilityLevel;
    @Column(name = "language", length = 32)
    private String language;
    @Column(name = "offline_cache", nullable = false)
    private boolean offlineCache;
    @Column(name = "theme", nullable = false, length = 16)
    private String theme;

    protected WorkbookUserStateEntity() {}

    public WorkbookUserStateEntity(String unitId, String subject, boolean favorite, Instant lastOpenedAt, Instant updatedAt) {
        this(unitId, subject, favorite, lastOpenedAt, true, true, "local", "standard", null, true, "system", updatedAt);
    }

    public WorkbookUserStateEntity(String unitId, String subject, boolean favorite, Instant lastOpenedAt,
                                   boolean autoSave, boolean autoSync, String defaultCreateLocation,
                                   String importCompatibilityLevel, String language, boolean offlineCache,
                                   String theme, Instant updatedAt) {
        this.id = new Id(unitId, subject);
        this.favorite = favorite;
        this.lastOpenedAt = lastOpenedAt;
        this.autoSave = autoSave;
        this.autoSync = autoSync;
        this.defaultCreateLocation = defaultCreateLocation == null ? "remote" : defaultCreateLocation;
        this.importCompatibilityLevel = importCompatibilityLevel == null ? "standard" : importCompatibilityLevel;
        this.language = language;
        this.offlineCache = offlineCache;
        this.theme = theme == null ? "system" : theme;
        this.updatedAt = updatedAt;
    }

    public Id getId() { return id; }
    public boolean isFavorite() { return favorite; }
    public Instant getLastOpenedAt() { return lastOpenedAt; }
    public Instant getUpdatedAt() { return updatedAt; }
    public boolean isAutoSave() { return autoSave; }
    public boolean isAutoSync() { return autoSync; }
    public String getDefaultCreateLocation() { return defaultCreateLocation; }
    public String getImportCompatibilityLevel() { return importCompatibilityLevel; }
    public String getLanguage() { return language; }
    public boolean isOfflineCache() { return offlineCache; }
    public String getTheme() { return theme; }

    public void update(Boolean favorite, Instant lastOpenedAt, Boolean autoSave, Boolean autoSync,
                       String defaultCreateLocation, String importCompatibilityLevel, String language,
                       Boolean offlineCache, String theme, Instant updatedAt) {
        if (favorite != null) this.favorite = favorite;
        if (lastOpenedAt != null) this.lastOpenedAt = lastOpenedAt;
        if (autoSave != null) this.autoSave = autoSave;
        if (autoSync != null) this.autoSync = autoSync;
        if (defaultCreateLocation != null) this.defaultCreateLocation = defaultCreateLocation;
        if (importCompatibilityLevel != null) this.importCompatibilityLevel = importCompatibilityLevel;
        if (language != null) this.language = language;
        if (offlineCache != null) this.offlineCache = offlineCache;
        if (theme != null) this.theme = theme;
        this.updatedAt = updatedAt;
    }

    @Embeddable
    public static class Id implements Serializable {
        @Column(name = "unit_id", nullable = false, length = 200)
        private String unitId;
        @Column(name = "subject", nullable = false, length = 500)
        private String subject;
        protected Id() {}
        public Id(String unitId, String subject) { this.unitId = unitId; this.subject = subject; }
        public String getUnitId() { return unitId; }
        public String getSubject() { return subject; }
        @Override public boolean equals(Object o) { return this == o || (o instanceof Id other && Objects.equals(unitId, other.unitId) && Objects.equals(subject, other.subject)); }
        @Override public int hashCode() { return Objects.hash(unitId, subject); }
    }
}
