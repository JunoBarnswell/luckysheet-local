package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import java.time.Instant;
import java.util.UUID;

public record AuditRecord(UUID auditId, String operationId, String unitId, String actorId, String eventType, String outcome, String reason, JsonNode details, Instant occurredAt) {
}
