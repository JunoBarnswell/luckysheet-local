package com.xc.luckysheet.server.contract;

import java.time.Instant;
public record AclEntry(String unitId, String subject, WorkbookRole role, Instant createdAt, Instant updatedAt) {
}
