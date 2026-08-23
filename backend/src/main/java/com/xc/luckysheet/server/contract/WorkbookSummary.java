package com.xc.luckysheet.server.contract;

import java.time.Instant;
public record WorkbookSummary(String unitId, String name, long revision, Instant updatedAt) {
}
