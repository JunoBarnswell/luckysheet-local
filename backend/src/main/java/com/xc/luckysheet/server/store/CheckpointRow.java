package com.xc.luckysheet.server.store;

import java.time.Instant;
public record CheckpointRow(String unitId, long revision, String snapshotJson, String checksum, Instant createdAt) {
}
