package com.xc.luckysheet.server.persistence;

import com.xc.luckysheet.server.contract.RangeAccessLevel;
import com.xc.luckysheet.server.contract.RangeRef;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.EnumType;
import jakarta.persistence.Enumerated;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;

import java.time.Instant;

@Entity
@Table(name = "workbook_access_region", indexes = {
        @Index(name = "workbook_access_region_sheet_idx", columnList = "unit_id,sheet_id,start_row,start_column")
})
public class RangeAccessRegionEntity {
    @Id
    @Column(name = "region_id", nullable = false, length = 200)
    private String id;

    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;

    @Column(name = "sheet_id", nullable = false, length = 200)
    private String sheetId;

    @Column(name = "start_row", nullable = false)
    private int startRow;

    @Column(name = "end_row", nullable = false)
    private int endRow;

    @Column(name = "start_column", nullable = false)
    private int startColumn;

    @Column(name = "end_column", nullable = false)
    private int endColumn;

    @Enumerated(EnumType.STRING)
    @Column(name = "default_access", nullable = false, length = 16)
    private RangeAccessLevel defaultAccess;

    @Column(name = "created_by", nullable = false, length = 500)
    private String createdBy;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected RangeAccessRegionEntity() {
    }

    public RangeAccessRegionEntity(String id, String unitId, String sheetId, RangeRef range,
                                   RangeAccessLevel defaultAccess, String createdBy, Instant now) {
        this.id = id;
        this.unitId = unitId;
        this.sheetId = sheetId;
        setRange(range);
        this.defaultAccess = defaultAccess;
        this.createdBy = createdBy;
        this.createdAt = now;
        this.updatedAt = now;
    }

    public void update(String sheetId, RangeRef range, RangeAccessLevel defaultAccess, Instant now) {
        this.sheetId = sheetId;
        setRange(range);
        this.defaultAccess = defaultAccess;
        this.updatedAt = now;
    }

    private void setRange(RangeRef range) {
        this.startRow = range.startRow();
        this.endRow = range.endRow();
        this.startColumn = range.startColumn();
        this.endColumn = range.endColumn();
    }

    public String getId() { return id; }
    public String getUnitId() { return unitId; }
    public String getSheetId() { return sheetId; }
    public RangeRef getRange() { return new RangeRef(sheetId, startRow, endRow, startColumn, endColumn); }
    public RangeAccessLevel getDefaultAccess() { return defaultAccess; }
    public String getCreatedBy() { return createdBy; }
    public Instant getCreatedAt() { return createdAt; }
    public Instant getUpdatedAt() { return updatedAt; }
}
