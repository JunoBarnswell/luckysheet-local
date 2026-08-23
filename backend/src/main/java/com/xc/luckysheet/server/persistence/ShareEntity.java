package com.xc.luckysheet.server.persistence;

import com.xc.luckysheet.server.contract.WorkbookAclRole;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.EnumType;
import jakarta.persistence.Enumerated;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.time.Instant;
import java.util.UUID;

@Entity
@Table(name = "workbook_share", indexes = {
        @Index(name = "workbook_share_unit_idx", columnList = "unit_id,expires_at"),
        @Index(name = "workbook_share_active_token_idx", columnList = "token_hash,revoked_at,expires_at")
})
public class ShareEntity {
    @Id
    @JdbcTypeCode(SqlTypes.VARCHAR)
    @Column(name = "share_id", nullable = false)
    private UUID shareId;

    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;

    @Column(name = "token_hash", nullable = false, unique = true, length = 128)
    private String tokenHash;

    @Enumerated(EnumType.STRING)
    @Column(name = "role", nullable = false, length = 16)
    private WorkbookAclRole role;

    @Column(name = "expires_at", nullable = false)
    private Instant expiresAt;

    @Column(name = "revoked_at")
    private Instant revokedAt;

    @Column(name = "created_by", nullable = false, length = 500)
    private String createdBy;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    protected ShareEntity() {
    }

    public ShareEntity(UUID shareId, String unitId, String tokenHash, WorkbookAclRole role, Instant expiresAt,
                       Instant revokedAt, String createdBy, Instant createdAt) {
        this.shareId = shareId;
        this.unitId = unitId;
        this.tokenHash = tokenHash;
        this.role = role;
        this.expiresAt = expiresAt;
        this.revokedAt = revokedAt;
        this.createdBy = createdBy;
        this.createdAt = createdAt;
    }

    public UUID getShareId() { return shareId; }
    public String getUnitId() { return unitId; }
    public String getTokenHash() { return tokenHash; }
    public WorkbookAclRole getRole() { return role; }
    public Instant getExpiresAt() { return expiresAt; }
    public Instant getRevokedAt() { return revokedAt; }
    public String getCreatedBy() { return createdBy; }
    public Instant getCreatedAt() { return createdAt; }

    public void revoke(Instant revokedAt) {
        this.revokedAt = revokedAt;
    }
}
