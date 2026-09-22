package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;

import java.time.Instant;

/** Persistent local account used when the server runs in local auth mode. */
@Entity
@Table(name = "app_user", indexes = {
        @Index(name = "app_user_username_idx", columnList = "username", unique = true),
        @Index(name = "app_user_enabled_idx", columnList = "enabled")
})
public class LocalUserEntity {
    @Id
    @Column(name = "user_id", nullable = false, length = 36)
    private String userId;

    @Column(name = "username", nullable = false, length = 200, unique = true)
    private String username;

    @Column(name = "password_hash", nullable = false, length = 255)
    private String passwordHash;

    @Column(name = "display_name", nullable = false, length = 200)
    private String displayName;

    @Column(name = "enabled", nullable = false)
    private boolean enabled;

    @Column(name = "admin", nullable = false)
    private boolean admin;

    @Column(name = "credential_version", nullable = false)
    private long credentialVersion;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected LocalUserEntity() {
    }

    public LocalUserEntity(String userId, String username, String passwordHash, String displayName,
                           boolean enabled, boolean admin, Instant createdAt, Instant updatedAt) {
        this.userId = userId;
        this.username = username;
        this.passwordHash = passwordHash;
        this.displayName = displayName;
        this.enabled = enabled;
        this.admin = admin;
        this.credentialVersion = 0L;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
    }

    public String getUserId() {
        return userId;
    }

    public String getUsername() {
        return username;
    }

    public String getPasswordHash() {
        return passwordHash;
    }

    public String getDisplayName() {
        return displayName;
    }

    public boolean isEnabled() {
        return enabled;
    }

    public boolean isAdmin() {
        return admin;
    }

    public long getCredentialVersion() {
        return credentialVersion;
    }

    public Instant getCreatedAt() {
        return createdAt;
    }

    public Instant getUpdatedAt() {
        return updatedAt;
    }

    public void setEnabled(boolean enabled) {
        this.enabled = enabled;
        this.updatedAt = Instant.now();
    }

    public void resetPassword(String passwordHash) {
        this.passwordHash = passwordHash;
        this.credentialVersion++;
        this.updatedAt = Instant.now();
    }
}
