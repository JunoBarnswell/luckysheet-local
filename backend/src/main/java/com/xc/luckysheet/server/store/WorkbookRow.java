package com.xc.luckysheet.server.store;

import java.time.Instant;
public record WorkbookRow(String unitId, String name, String snapshotJson, long snapshotRevision, long revision, Instant createdAt, Instant updatedAt) {
}
