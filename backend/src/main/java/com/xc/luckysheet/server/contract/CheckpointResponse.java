package com.xc.luckysheet.server.contract;

public record CheckpointResponse(String unitId, long revision, String checksum, boolean created) {
}
