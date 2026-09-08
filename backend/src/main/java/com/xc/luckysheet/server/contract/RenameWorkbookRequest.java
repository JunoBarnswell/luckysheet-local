package com.xc.luckysheet.server.contract;

/** A workbook rename is a kernel mutation, never a catalog metadata patch. */
public record RenameWorkbookRequest(String name) {
    public RenameWorkbookRequest {
        if (name == null || name.isBlank() || name.trim().length() > GeneratedWorkbookContract.MAX_WORKBOOK_NAME_LENGTH) {
            throw new IllegalArgumentException("A valid workbook name is required");
        }
        name = name.trim();
    }
}
