package com.xc.luckysheet.server.persistence;

import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.contract.WorkbookSource;
import com.xc.luckysheet.server.contract.WorkbookStorageLocation;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.EnumType;
import jakarta.persistence.Enumerated;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;
import jakarta.persistence.Version;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.time.Instant;

/** Canonical workbook state. JSON is deliberately stored as portable text. */
@Entity
@Table(name = "workbooks", indexes = {
        @Index(name = "workbooks_owner_updated_idx", columnList = "owner_subject,updated_at"),
        @Index(name = "workbooks_space_folder_idx", columnList = "space_id,folder_id,deleted_at")
})
public class WorkbookEntity {
    @Id
    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;

    @Column(name = "name", nullable = false, length = 255)
    private String name;

    @JdbcTypeCode(SqlTypes.LONGVARCHAR)
    @Column(name = "snapshot_json", nullable = false)
    private String snapshotJson;

    @Column(name = "snapshot_revision", nullable = false)
    private long snapshotRevision;

    @Column(name = "revision", nullable = false)
    private long revision;

    /** Optimistic backstop for catalog/lifecycle writes outside operation replay. */
    @Version
    @Column(name = "entity_version", nullable = false)
    private long entityVersion;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    @Column(name = "owner_subject", nullable = false, length = 500)
    private String ownerSubject;

    @Column(name = "space_id", length = 200)
    private String spaceId;

    @Column(name = "folder_id", length = 200)
    private String folderId;

    @Enumerated(EnumType.STRING)
    @Column(name = "storage_location", nullable = false, length = 16)
    private WorkbookStorageLocation storageLocation;

    @Enumerated(EnumType.STRING)
    @Column(name = "source", nullable = false, length = 32)
    private WorkbookSource source;

    @Enumerated(EnumType.STRING)
    @Column(name = "lifecycle", nullable = false, length = 16)
    private WorkbookLifecycle lifecycle;

    @Column(name = "deleted_at")
    private Instant deletedAt;

    protected WorkbookEntity() {
    }

    public WorkbookEntity(String unitId, String name, String snapshotJson, long snapshotRevision, long revision,
                          Instant createdAt, Instant updatedAt) {
        this(unitId, name, snapshotJson, snapshotRevision, revision, createdAt, updatedAt, "", null, null,
                WorkbookStorageLocation.REMOTE, WorkbookSource.NATIVE, WorkbookLifecycle.ACTIVE, null);
    }

    public WorkbookEntity(String unitId, String name, String snapshotJson, long snapshotRevision, long revision,
                          Instant createdAt, Instant updatedAt, String ownerSubject, String spaceId, String folderId,
                          WorkbookStorageLocation storageLocation, WorkbookSource source, WorkbookLifecycle lifecycle,
                          Instant deletedAt) {
        this.unitId = unitId;
        this.name = name;
        this.snapshotJson = snapshotJson;
        this.snapshotRevision = snapshotRevision;
        this.revision = revision;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
        this.ownerSubject = ownerSubject == null ? "" : ownerSubject;
        this.spaceId = spaceId;
        this.folderId = folderId;
        this.storageLocation = storageLocation == null ? WorkbookStorageLocation.REMOTE : storageLocation;
        this.source = source == null ? WorkbookSource.NATIVE : source;
        this.lifecycle = lifecycle == null ? WorkbookLifecycle.ACTIVE : lifecycle;
        this.deletedAt = deletedAt;
    }

    public String getUnitId() {
        return unitId;
    }

    public String getName() {
        return name;
    }

    public String getSnapshotJson() {
        return snapshotJson;
    }

    public long getSnapshotRevision() {
        return snapshotRevision;
    }

    public long getRevision() {
        return revision;
    }

    public Instant getCreatedAt() {
        return createdAt;
    }

    public Instant getUpdatedAt() {
        return updatedAt;
    }

    public String getOwnerSubject() { return ownerSubject; }
    public String getSpaceId() { return spaceId; }
    public String getFolderId() { return folderId; }
    public WorkbookStorageLocation getStorageLocation() { return storageLocation; }
    public WorkbookSource getSource() { return source; }
    public WorkbookLifecycle getLifecycle() { return lifecycle; }
    public Instant getDeletedAt() { return deletedAt; }

    public void updateRevision(long revision, Instant updatedAt) {
        this.revision = revision;
        this.updatedAt = updatedAt;
    }

    public void updateRevisionAndName(long revision, String name, Instant updatedAt) {
        this.revision = revision;
        if (name != null && !name.isBlank()) this.name = name;
        this.updatedAt = updatedAt;
    }

    public void updateSnapshot(long revision, String snapshotJson, long snapshotRevision, Instant updatedAt) {
        this.revision = revision;
        this.snapshotJson = snapshotJson;
        this.snapshotRevision = snapshotRevision;
        this.updatedAt = updatedAt;
    }

    public void updateLocation(String spaceId, String folderId, Instant updatedAt) {
        this.spaceId = spaceId;
        this.folderId = folderId;
        this.updatedAt = updatedAt;
    }

    public void moveToTrash(Instant deletedAt) {
        this.lifecycle = WorkbookLifecycle.TRASHED;
        this.deletedAt = deletedAt;
        this.updatedAt = deletedAt;
    }

    public void restoreFromTrash(Instant restoredAt) {
        this.lifecycle = WorkbookLifecycle.ACTIVE;
        this.deletedAt = null;
        this.updatedAt = restoredAt;
    }
}
