package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;

@Entity
@Table(name = "workbook_access_revision")
public class WorkbookAccessRevisionEntity {
    @Id
    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;

    @Column(name = "revision", nullable = false)
    private long revision;

    protected WorkbookAccessRevisionEntity() {
    }

    public WorkbookAccessRevisionEntity(String unitId, long revision) {
        this.unitId = unitId;
        this.revision = revision;
    }

    public String getUnitId() { return unitId; }
    public long getRevision() { return revision; }
    public long increment() { return ++revision; }
}
