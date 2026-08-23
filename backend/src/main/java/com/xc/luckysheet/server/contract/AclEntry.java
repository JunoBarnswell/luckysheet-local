package com.xc.luckysheet.server.contract;

import java.time.Instant;
public record AclEntry(String unitId, String subject, WorkbookAclRole role, Instant createdAt, Instant updatedAt) {
}
