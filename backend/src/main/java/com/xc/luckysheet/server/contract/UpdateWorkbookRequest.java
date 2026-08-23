package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

/** Metadata patch deliberately distinguishes an omitted field from JSON null. */
public final class UpdateWorkbookRequest {
    private final boolean spaceIdSpecified;
    private final boolean folderIdSpecified;
    private final String spaceId;
    private final String folderId;

    @JsonCreator
    public UpdateWorkbookRequest(
            @JsonProperty("spaceId") JsonNode spaceId,
            @JsonProperty("folderId") JsonNode folderId
    ) {
        this.spaceIdSpecified = spaceId != null;
        this.folderIdSpecified = folderId != null;
        if (!spaceIdSpecified && !folderIdSpecified) {
            throw new IllegalArgumentException("At least one workbook location field is required");
        }
        this.spaceId = nullableIdentity(spaceId, "spaceId");
        this.folderId = nullableIdentity(folderId, "folderId");
    }

    public boolean spaceIdSpecified() { return spaceIdSpecified; }
    public boolean folderIdSpecified() { return folderIdSpecified; }
    public String spaceId() { return spaceId; }
    public String folderId() { return folderId; }

    private static String nullableIdentity(JsonNode value, String label) {
        if (value == null || value.isNull()) return null;
        if (!value.isTextual()) throw new IllegalArgumentException(label + " must be a string or null");
        String normalized = value.asText().trim();
        if (normalized.length() > 200) throw new IllegalArgumentException(label + " is too long");
        return normalized.isEmpty() ? null : normalized;
    }
}
