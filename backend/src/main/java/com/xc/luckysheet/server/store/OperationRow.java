package com.xc.luckysheet.server.store;

import java.time.Instant;
public record OperationRow(String operationId, String unitId, long revision, String actorSubject, long clientSequence, long baseRevision, String envelopeJson, Instant committedAt) {
}
