package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;

public record CreateWorkbookRequest(
        @JsonProperty("unitId") String unitId,
        @JsonProperty("name") String name,
        @JsonProperty("snapshot") JsonNode snapshot,
        @JsonProperty("spaceId") String spaceId,
        @JsonProperty("folderId") String folderId,
        @JsonProperty("source") WorkbookSource source
) {
    public CreateWorkbookRequest(String unitId, String name, JsonNode snapshot) {
        this(unitId, name, snapshot, null, null, WorkbookSource.NATIVE);
    }

    @JsonCreator
    public CreateWorkbookRequest {
        if (unitId == null || unitId.isBlank() || name == null || name.isBlank() || snapshot == null || !snapshot.isObject()) {
            throw new IllegalArgumentException("unitId, name and object snapshot are required");
        }
        if (unitId.length() > 200 || name.length() > GeneratedWorkbookContract.MAX_WORKBOOK_NAME_LENGTH) {
            throw new IllegalArgumentException("Workbook identity is too long");
        }
        if (source == null) source = WorkbookSource.NATIVE;
    }
}
