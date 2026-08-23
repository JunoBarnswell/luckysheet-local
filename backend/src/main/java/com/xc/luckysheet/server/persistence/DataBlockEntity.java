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
@Table(name = "workbook_data_block", indexes = @Index(name = "workbook_data_block_lookup_idx", columnList = "unit_id,source_id,checksum"))
public class DataBlockEntity {
    @EmbeddedId
    private Id id;

    @Column(name = "checksum", nullable = false, length = 64)
    private String checksum;

    @Column(name = "byte_length", nullable = false)
    private int byteLength;

    @JdbcTypeCode(SqlTypes.LONGVARBINARY)
    @Column(name = "content", nullable = false)
    private byte[] content;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected DataBlockEntity() {
    }

    public DataBlockEntity(String unitId, String sourceId, String blockId, String checksum, int byteLength,
                           byte[] content, Instant createdAt, Instant updatedAt) {
        this.id = new Id(unitId, sourceId, blockId);
        this.checksum = checksum;
        this.byteLength = byteLength;
        this.content = content;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
    }

    public Id getId() { return id; }
    public String getChecksum() { return checksum; }
    public int getByteLength() { return byteLength; }
    public byte[] getContent() { return content; }
    public Instant getCreatedAt() { return createdAt; }
    public Instant getUpdatedAt() { return updatedAt; }

    public void update(String checksum, int byteLength, byte[] content, Instant updatedAt) {
        this.checksum = checksum;
        this.byteLength = byteLength;
        this.content = content;
        this.updatedAt = updatedAt;
    }

    @Embeddable
    public static class Id implements Serializable {
        @Column(name = "unit_id", nullable = false, length = 200)
        private String unitId;

        @Column(name = "source_id", nullable = false, length = 200)
        private String sourceId;

        @Column(name = "block_id", nullable = false, length = 200)
        private String blockId;

        protected Id() {
        }

        public Id(String unitId, String sourceId, String blockId) {
            this.unitId = unitId;
            this.sourceId = sourceId;
            this.blockId = blockId;
        }

        public String getUnitId() { return unitId; }
        public String getSourceId() { return sourceId; }
        public String getBlockId() { return blockId; }

        @Override
        public boolean equals(Object object) {
            if (this == object) return true;
            if (!(object instanceof Id other)) return false;
            return Objects.equals(unitId, other.unitId)
                    && Objects.equals(sourceId, other.sourceId)
                    && Objects.equals(blockId, other.blockId);
        }

        @Override
        public int hashCode() {
            return Objects.hash(unitId, sourceId, blockId);
        }
    }
}
