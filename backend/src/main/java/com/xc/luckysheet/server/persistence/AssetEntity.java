package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Embeddable;
import jakarta.persistence.EmbeddedId;
import jakarta.persistence.Entity;
import jakarta.persistence.Table;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.io.Serializable;
import java.time.Instant;
import java.util.Objects;

@Entity
@Table(name = "workbook_asset")
public class AssetEntity {
    @EmbeddedId
    private Id id;

    @Column(name = "content_hash", nullable = false, length = 64)
    private String contentHash;

    @Column(name = "mime_type", nullable = false, length = 127)
    private String mimeType;

    @Column(name = "byte_length", nullable = false)
    private int byteLength;

    @Column(name = "width")
    private Integer width;

    @Column(name = "height")
    private Integer height;

    @JdbcTypeCode(SqlTypes.LONGVARBINARY)
    @Column(name = "content", nullable = false)
    private byte[] content;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected AssetEntity() {
    }

    public AssetEntity(String unitId, String assetId, String contentHash, String mimeType, int byteLength,
                       Integer width, Integer height, byte[] content, Instant createdAt, Instant updatedAt) {
        this.id = new Id(unitId, assetId);
        this.contentHash = contentHash;
        this.mimeType = mimeType;
        this.byteLength = byteLength;
        this.width = width;
        this.height = height;
        this.content = content;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
    }

    public Id getId() { return id; }
    public String getContentHash() { return contentHash; }
    public String getMimeType() { return mimeType; }
    public int getByteLength() { return byteLength; }
    public Integer getWidth() { return width; }
    public Integer getHeight() { return height; }
    public byte[] getContent() { return content; }
    public Instant getUpdatedAt() { return updatedAt; }

    @Embeddable
    public static class Id implements Serializable {
        @Column(name = "unit_id", nullable = false, length = 200)
        private String unitId;

        @Column(name = "asset_id", nullable = false, length = 200)
        private String assetId;

        protected Id() {
        }

        public Id(String unitId, String assetId) {
            this.unitId = unitId;
            this.assetId = assetId;
        }

        public String getUnitId() { return unitId; }
        public String getAssetId() { return assetId; }

        @Override
        public boolean equals(Object object) {
            if (this == object) return true;
            if (!(object instanceof Id other)) return false;
            return Objects.equals(unitId, other.unitId) && Objects.equals(assetId, other.assetId);
        }

        @Override
        public int hashCode() { return Objects.hash(unitId, assetId); }
    }
}
