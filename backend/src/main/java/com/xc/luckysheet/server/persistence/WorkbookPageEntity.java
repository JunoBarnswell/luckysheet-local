package com.xc.luckysheet.server.persistence;

import jakarta.persistence.*;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;
import java.time.Instant;

/** Content-addressed immutable native bytes. Coordinates and source revision belong to manifest descriptors. */
@Entity
@Table(name = "workbook_pages", uniqueConstraints = @UniqueConstraint(name = "workbook_page_content_idx", columnNames = {"unit_id", "checksum"}))
public class WorkbookPageEntity {
    @Id @Column(name = "page_id", nullable = false, length = 64) private String pageId;
    @Column(name = "unit_id", nullable = false, length = 200) private String unitId;
    @Column(name = "checksum", nullable = false, length = 64) private String checksum;
    @Column(name = "byte_length", nullable = false) private long byteLength;
    @JdbcTypeCode(SqlTypes.LONGVARCHAR) @Column(name = "payload_base64", nullable = false) private String payloadBase64;
    @Column(name = "created_at", nullable = false) private Instant createdAt;
    protected WorkbookPageEntity() { }
    public WorkbookPageEntity(String pageId, String unitId, String checksum, long byteLength, String payloadBase64, Instant createdAt) {
        this.pageId = pageId; this.unitId = unitId; this.checksum = checksum;
        this.byteLength = byteLength; this.payloadBase64 = payloadBase64; this.createdAt = createdAt;
    }
    public String getPageId() { return pageId; }
    public String getUnitId() { return unitId; }
    public String getChecksum() { return checksum; }
    public long getByteLength() { return byteLength; }
    public String getPayloadBase64() { return payloadBase64; }
}
