package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;

import java.time.Instant;
import java.util.UUID;

@JsonInclude(JsonInclude.Include.NON_NULL)
public record ShareResponse(
        @JsonProperty("shareId") UUID shareId,
        @JsonProperty("unitId") String unitId,
        @JsonProperty("role") WorkbookAclRole role,
        @JsonProperty("expiresAt") Instant expiresAt,
        @JsonProperty("revokedAt") Instant revokedAt,
        @JsonProperty("createdBy") String createdBy,
        @JsonProperty("createdAt") Instant createdAt,
        @JsonProperty("token") String token
) {
    public static ShareResponse listed(com.xc.luckysheet.server.store.ShareRow row) {
        return new ShareResponse(row.shareId(), row.unitId(), row.role(), row.expiresAt(), row.revokedAt(), row.createdBy(), row.createdAt(), null);
    }
}
