package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.time.Instant;

public record ShareCreateRequest(
        @JsonProperty("role") String role,
        @JsonProperty("expiresAt") Instant expiresAt
) {
    @JsonCreator
    public ShareCreateRequest {
        if (role == null || role.isBlank()) throw new IllegalArgumentException("Share role is required");
    }
}
