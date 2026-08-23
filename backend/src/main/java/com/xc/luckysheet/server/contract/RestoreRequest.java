package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

public record RestoreRequest(@JsonProperty("targetRevision") long targetRevision, @JsonProperty("reason") String reason) {
    @JsonCreator
    public RestoreRequest {
        if (targetRevision < 0) throw new IllegalArgumentException("targetRevision must be non-negative");
        if (reason == null || reason.isBlank()) throw new IllegalArgumentException("reason is required");
    }
}
