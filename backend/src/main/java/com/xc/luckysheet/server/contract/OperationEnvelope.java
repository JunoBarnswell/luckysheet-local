package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.time.Instant;
import java.util.List;
import java.util.Objects;

public record OperationEnvelope(
        @JsonProperty("schema") String schema,
        @JsonProperty("operationId") String operationId,
        @JsonProperty("unitId") String unitId,
        @JsonProperty("clientSequence") long clientSequence,
        @JsonProperty("baseRevision") long baseRevision,
        @JsonProperty("mutations") List<OperationMutation> mutations,
        @JsonProperty("createdAt") Instant createdAt
) {
    public static final String SCHEMA = "OperationEnvelope";

    @JsonCreator
    public OperationEnvelope {
        if (!SCHEMA.equals(schema)) throw new IllegalArgumentException("schema must be OperationEnvelope");
        if (operationId == null || operationId.isBlank()) throw new IllegalArgumentException("operationId is required");
        if (unitId == null || unitId.isBlank()) throw new IllegalArgumentException("unitId is required");
        if (clientSequence < 1) throw new IllegalArgumentException("clientSequence must be positive");
        if (baseRevision < 0) throw new IllegalArgumentException("baseRevision must be non-negative");
        if (mutations == null || mutations.isEmpty()) throw new IllegalArgumentException("mutations must not be empty");
        mutations = List.copyOf(mutations);
        Objects.requireNonNull(createdAt, "createdAt is required");
    }
}
