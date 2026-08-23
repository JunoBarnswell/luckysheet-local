package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.time.Instant;

/** Canonical workbook state. JSON is deliberately stored as portable text. */
@Entity
@Table(name = "workbooks")
public class WorkbookEntity {
    @Id
    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;

    @Column(name = "name", nullable = false, length = 500)
    private String name;

    @JdbcTypeCode(SqlTypes.LONGVARCHAR)
    @Column(name = "snapshot_json", nullable = false)
    private String snapshotJson;

    @Column(name = "snapshot_revision", nullable = false)
    private long snapshotRevision;

    @Column(name = "revision", nullable = false)
    private long revision;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected WorkbookEntity() {
    }

    public WorkbookEntity(String unitId, String name, String snapshotJson, long snapshotRevision, long revision,
                          Instant createdAt, Instant updatedAt) {
        this.unitId = unitId;
        this.name = name;
        this.snapshotJson = snapshotJson;
        this.snapshotRevision = snapshotRevision;
        this.revision = revision;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
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

    public void updateRevision(long revision, Instant updatedAt) {
        this.revision = revision;
        this.updatedAt = updatedAt;
    }

    public void updateSnapshot(long revision, String snapshotJson, long snapshotRevision, Instant updatedAt) {
        this.revision = revision;
        this.snapshotJson = snapshotJson;
        this.snapshotRevision = snapshotRevision;
        this.updatedAt = updatedAt;
    }
}
