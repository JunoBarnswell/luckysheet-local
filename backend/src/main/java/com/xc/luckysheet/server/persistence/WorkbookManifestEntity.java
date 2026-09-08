package com.xc.luckysheet.server.persistence;

import jakarta.persistence.*;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;
import java.time.Instant;

/** Immutable v11 manifest checkpoint. The workbook row owns the current revision pointer. */
@Entity
@Table(name = "workbook_manifests", uniqueConstraints = @UniqueConstraint(name = "workbook_manifest_revision_idx", columnNames = {"unit_id", "revision"}))
public class WorkbookManifestEntity {
    @Id @Column(name = "manifest_id", nullable = false, length = 64) private String manifestId;
    @Column(name = "unit_id", nullable = false, length = 200) private String unitId;
    @Column(name = "revision", nullable = false) private long revision;
    @Column(name = "manifest_version", nullable = false) private int manifestVersion;
    @JdbcTypeCode(SqlTypes.LONGVARCHAR) @Column(name = "manifest_json", nullable = false) private String manifestJson;
    @Column(name = "checksum", nullable = false, length = 64) private String checksum;
    @Column(name = "created_at", nullable = false) private Instant createdAt;
    protected WorkbookManifestEntity() { }
    public WorkbookManifestEntity(String manifestId, String unitId, long revision, String manifestJson, String checksum, Instant createdAt) {
        this.manifestId = manifestId; this.unitId = unitId; this.revision = revision; this.manifestVersion = 11;
        this.manifestJson = manifestJson; this.checksum = checksum; this.createdAt = createdAt;
    }
    public String getManifestId() { return manifestId; }
    public String getUnitId() { return unitId; }
    public long getRevision() { return revision; }
    public int getManifestVersion() { return manifestVersion; }
    public String getManifestJson() { return manifestJson; }
    public String getChecksum() { return checksum; }
}
