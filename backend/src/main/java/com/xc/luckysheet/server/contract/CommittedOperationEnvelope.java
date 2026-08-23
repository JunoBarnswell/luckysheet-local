package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;
import java.time.Instant;
import java.util.List;

public record CommittedOperationEnvelope(
        @JsonProperty("schema") String schema,
        @JsonProperty("operationId") String operationId,
        @JsonProperty("unitId") String unitId,
        @JsonProperty("actorId") String actorId,
        @JsonProperty("clientSequence") long clientSequence,
        @JsonProperty("baseRevision") long baseRevision,
        @JsonProperty("revision") long revision,
        @JsonProperty("mutations") List<CommittedOperationMutation> mutations,
        @JsonProperty("createdAt") Instant createdAt,
        @JsonProperty("committedAt") Instant committedAt
) {
    @JsonCreator
    public CommittedOperationEnvelope {
        if (!OperationEnvelope.SCHEMA.equals(schema)) throw new IllegalArgumentException("schema must be OperationEnvelope");
        if (operationId == null || operationId.isBlank() || unitId == null || unitId.isBlank() || actorId == null || actorId.isBlank()) throw new IllegalArgumentException("committed operation identity is required");
        if (clientSequence < 1 || baseRevision < 0 || revision < 1) throw new IllegalArgumentException("Invalid committed operation revision");
        if (mutations == null || mutations.isEmpty()) throw new IllegalArgumentException("mutations must not be empty");
        mutations = List.copyOf(mutations);
        if (createdAt == null || committedAt == null) throw new IllegalArgumentException("operation timestamps are required");
    }

    public static CommittedOperationEnvelope from(
            OperationEnvelope operation,
            String actorId,
            long revision,
            Instant committedAt,
            List<CommittedOperationMutation> mutations
    ) {
        return new CommittedOperationEnvelope(
                OperationEnvelope.SCHEMA,
                operation.operationId(),
                operation.unitId(),
                actorId,
                operation.clientSequence(),
                operation.baseRevision(),
                revision,
                mutations,
                // Client clocks are neither trusted nor replay authority.
                // The persisted envelope has one server-issued event time.
                committedAt,
                committedAt
        );
    }
}
