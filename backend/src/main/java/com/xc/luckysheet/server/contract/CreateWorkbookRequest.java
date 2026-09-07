package com.xc.luckysheet.server.contract;
import com.fasterxml.jackson.databind.JsonNode;
/** Creation intent. A browser-authored snapshot is never an import authority. */
public record CreateWorkbookRequest(String unitId, String name, JsonNode sheets,
                                    String spaceId, String folderId, WorkbookSource source) {
    public CreateWorkbookRequest {
        if (unitId == null || unitId.isBlank() || unitId.length() > 200 || name == null || name.isBlank()
                || name.length() > GeneratedWorkbookContract.MAX_WORKBOOK_NAME_LENGTH)
            throw new IllegalArgumentException("A valid unitId and name are required");
        if (sheets != null && !sheets.isArray()) throw new IllegalArgumentException("sheets must be an array");
        if (source == null) source = WorkbookSource.NATIVE;
        if (source != WorkbookSource.NATIVE) throw new IllegalArgumentException("Document imports require the native import endpoint");
    }
}
