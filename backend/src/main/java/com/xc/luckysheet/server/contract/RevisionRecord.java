package com.xc.luckysheet.server.contract;

import java.time.Instant;

public record RevisionRecord(String operationId, long revision, Instant createdAt, CommittedOperationEnvelope payload) {
}
