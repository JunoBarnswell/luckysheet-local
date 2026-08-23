package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Embeddable;
import jakarta.persistence.EmbeddedId;
import jakarta.persistence.Entity;
import jakarta.persistence.Index;
import jakarta.persistence.Table;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.io.Serializable;
import java.time.Instant;
import java.util.Objects;

@Entity
@Table(name = "snapshot_checkpoint", indexes = @Index(name = "snapshot_checkpoint_unit_revision_idx", columnList = "unit_id,revision"))
public class CheckpointEntity {
    @EmbeddedId
    private Id id;

    @JdbcTypeCode(SqlTypes.LONGVARCHAR)
    @Column(name = "snapshot_json", nullable = false)
    private String snapshotJson;

    @Column(name = "checksum", nullable = false, length = 64)
    private String checksum;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    protected CheckpointEntity() {
    }

    public CheckpointEntity(String unitId, long revision, String snapshotJson, String checksum, Instant createdAt) {
        this.id = new Id(unitId, revision);
        this.snapshotJson = snapshotJson;
        this.checksum = checksum;
        this.createdAt = createdAt;
    }

    public Id getId() { return id; }
    public String getSnapshotJson() { return snapshotJson; }
    public String getChecksum() { return checksum; }
    public Instant getCreatedAt() { return createdAt; }

    public void update(String snapshotJson, String checksum, Instant createdAt) {
        this.snapshotJson = snapshotJson;
        this.checksum = checksum;
        this.createdAt = createdAt;
    }

    @Embeddable
    public static class Id implements Serializable {
        @Column(name = "unit_id", nullable = false, length = 200)
        private String unitId;

        @Column(name = "revision", nullable = false)
        private long revision;

        protected Id() {
        }

        public Id(String unitId, long revision) {
            this.unitId = unitId;
            this.revision = revision;
        }

        public String getUnitId() { return unitId; }
        public long getRevision() { return revision; }

        @Override
        public boolean equals(Object object) {
            if (this == object) return true;
            if (!(object instanceof Id other)) return false;
            return revision == other.revision && Objects.equals(unitId, other.unitId);
        }

        @Override
        public int hashCode() {
            return Objects.hash(unitId, revision);
        }
    }
}
