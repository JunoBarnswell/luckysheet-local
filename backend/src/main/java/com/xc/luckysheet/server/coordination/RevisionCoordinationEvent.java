package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;

import java.util.UUID;

/** Typed Redis revision notification whose routing identity must match its durable operation. */
public record RevisionCoordinationEvent(
        @JsonProperty("kind") String kind,
        @JsonProperty("eventId") String eventId,
        @JsonProperty("unitId") String unitId,
        @JsonProperty("operationId") String operationId,
        @JsonProperty("revision") long revision,
        @JsonProperty("operation") CommittedOperationEnvelope operation
) {
    @JsonCreator
    public RevisionCoordinationEvent {
        if (!"revision".equals(kind)) throw new IllegalArgumentException("COORDINATION_REVISION_KIND_INVALID");
        if (eventId == null || eventId.isBlank() || !UUID.fromString(eventId).toString().equals(eventId)) {
            throw new IllegalArgumentException("COORDINATION_REVISION_EVENT_ID_INVALID");
        }
        if (unitId == null || unitId.isBlank() || operationId == null || operationId.isBlank() || revision < 1
                || operation == null) {
            throw new IllegalArgumentException("COORDINATION_REVISION_IDENTITY_INCOMPLETE");
        }
        if (!unitId.equals(operation.unitId()) || !operationId.equals(operation.operationId())
                || revision != operation.revision()) {
            throw new IllegalArgumentException("COORDINATION_REVISION_IDENTITY_MISMATCH");
        }
    }
}
