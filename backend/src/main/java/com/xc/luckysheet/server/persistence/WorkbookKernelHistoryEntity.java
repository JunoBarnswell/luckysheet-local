package com.xc.luckysheet.server.persistence;

import jakarta.persistence.*;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;
import java.time.Instant;

/** Proven native history: immutable metadata and page references, never a second cell model. */
@Entity
@Table(name = "workbook_kernel_history", uniqueConstraints = @UniqueConstraint(name = "workbook_history_operation_idx", columnNames = {"unit_id", "operation_id"}))
public class WorkbookKernelHistoryEntity {
    @Id @Column(name = "history_id", nullable = false, length = 64) private String historyId;
    @Column(name = "unit_id", nullable = false, length = 200) private String unitId;
    @Column(name = "operation_id", nullable = false, length = 200) private String operationId;
    @Column(name = "base_revision", nullable = false) private long baseRevision;
    @Column(name = "revision", nullable = false) private long revision;
    @JdbcTypeCode(SqlTypes.LONGVARCHAR) @Column(name = "history_json", nullable = false) private String historyJson;
    @JdbcTypeCode(SqlTypes.LONGVARCHAR) @Column(name = "result_metadata_json", nullable = false) private String resultMetadataJson;
    @Column(name = "checksum", nullable = false, length = 64) private String checksum;
    @Column(name = "created_at", nullable = false) private Instant createdAt;
    protected WorkbookKernelHistoryEntity() { }
    public WorkbookKernelHistoryEntity(String historyId, String unitId, String operationId, long baseRevision, long revision, String historyJson, String checksum, String resultMetadataJson, Instant createdAt) {
        this.historyId = historyId; this.unitId = unitId; this.operationId = operationId; this.baseRevision = baseRevision;
        this.revision = revision; this.historyJson = historyJson; this.checksum = checksum; this.resultMetadataJson = resultMetadataJson; this.createdAt = createdAt;
    }
    public String getHistoryJson() { return historyJson; }
    public String getResultMetadataJson() { return resultMetadataJson; }
    public String getChecksum() { return checksum; }
    public long getBaseRevision() { return baseRevision; }
    public long getRevision() { return revision; }
}
