package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

/** Client intent metadata; the server remains the authority for target ownership and revision. */
public record OperationIntent(
        @JsonProperty("type") String type,
        @JsonProperty("targetOperationId") String targetOperationId,
        @JsonProperty("targetBaseRevision") long targetBaseRevision
) {
    public static final String UNDO = "undo";

    @JsonCreator
    public OperationIntent {
        if (!UNDO.equals(type)) throw new IllegalArgumentException("operation intent type must be undo");
        if (targetOperationId == null || targetOperationId.isBlank()) throw new IllegalArgumentException("undo target operationId is required");
        if (targetBaseRevision < 0) throw new IllegalArgumentException("undo targetBaseRevision must be non-negative");
    }
}
