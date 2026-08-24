package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.time.Instant;

@Entity
@Table(name = "workbook_source_artifact", indexes = @Index(name = "workbook_source_artifact_checksum_idx", columnList = "checksum"))
public class WorkbookSourceArtifactEntity {
    @Id
    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;

    @Column(name = "file_name", nullable = false, length = 500)
    private String fileName;

    @Column(name = "mime_type", nullable = false, length = 200)
    private String mimeType;

    @Column(name = "checksum", nullable = false, length = 64)
    private String checksum;

    @Column(name = "byte_length", nullable = false)
    private long byteLength;

    @JdbcTypeCode(SqlTypes.LONGVARBINARY)
    @Column(name = "content", nullable = false)
    private byte[] content;

    @JdbcTypeCode(SqlTypes.LONGVARCHAR)
    @Column(name = "detected_features_json", nullable = false)
    private String detectedFeaturesJson;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected WorkbookSourceArtifactEntity() {}

    public WorkbookSourceArtifactEntity(String unitId, String fileName, String mimeType, String checksum, long byteLength,
                                        byte[] content, String detectedFeaturesJson, Instant createdAt, Instant updatedAt) {
        this.unitId = unitId;
        this.fileName = fileName;
        this.mimeType = mimeType;
        this.checksum = checksum;
        this.byteLength = byteLength;
        this.content = content;
        this.detectedFeaturesJson = detectedFeaturesJson == null ? "[]" : detectedFeaturesJson;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
    }

    public String getUnitId() { return unitId; }
    public String getFileName() { return fileName; }
    public String getMimeType() { return mimeType; }
    public String getChecksum() { return checksum; }
    public long getByteLength() { return byteLength; }
    public byte[] getContent() { return content; }
    public String getDetectedFeaturesJson() { return detectedFeaturesJson; }
    public int getXlsxCodecVersion() {
        java.util.regex.Matcher matcher = java.util.regex.Pattern.compile("\\\"xlsxCodecVersion\\\"\\s*:\\s*(\\d+)").matcher(detectedFeaturesJson);
        return matcher.find() ? Integer.parseInt(matcher.group(1)) : 1;
    }
    public Instant getCreatedAt() { return createdAt; }
    public Instant getUpdatedAt() { return updatedAt; }

    public void update(String fileName, String mimeType, String checksum, long byteLength, byte[] content,
                       String detectedFeaturesJson, Instant updatedAt) {
        this.fileName = fileName;
        this.mimeType = mimeType;
        this.checksum = checksum;
        this.byteLength = byteLength;
        this.content = content;
        this.detectedFeaturesJson = detectedFeaturesJson == null ? "[]" : detectedFeaturesJson;
        this.updatedAt = updatedAt;
    }
}
