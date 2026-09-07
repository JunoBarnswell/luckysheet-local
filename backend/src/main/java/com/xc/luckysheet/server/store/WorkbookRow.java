package com.xc.luckysheet.server.store;

import java.time.Instant;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;

public record WorkbookRow(String unitId, String name, long revision,
                          WorkbookLifecycle lifecycle, Instant createdAt, Instant updatedAt) {
}
