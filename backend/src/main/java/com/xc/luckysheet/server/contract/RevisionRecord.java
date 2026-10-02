package com.xc.luckysheet.server.contract;

import java.time.Instant;

/** Revision metadata always survives projection; payload is omitted when current ACLs hide an affected range. */
public record RevisionRecord(
        String operationId,
        long revision,
        Instant createdAt,
        String actorId,
        long accessRevision,
        CommittedOperationEnvelope payload,
        boolean resyncRequired
) {
    public RevisionRecord {
        if (operationId == null || operationId.isBlank() || revision < 1 || createdAt == null
                || actorId == null || actorId.isBlank() || accessRevision < 0
                || (resyncRequired == (payload != null))) {
            throw new IllegalArgumentException("Revision record projection is invalid");
        }
        if (payload != null && (payload.revision() != revision || !payload.operationId().equals(operationId))) {
            throw new IllegalArgumentException("Revision record payload identity does not match metadata");
        }
    }
}
