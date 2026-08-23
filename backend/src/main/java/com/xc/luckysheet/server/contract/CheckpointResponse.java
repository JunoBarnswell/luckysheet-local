package com.xc.luckysheet.server.contract;

public record CheckpointResponse(WorkbookSnapshotResponse snapshot, boolean created) {
}
