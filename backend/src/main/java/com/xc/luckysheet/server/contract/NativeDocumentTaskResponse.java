package com.xc.luckysheet.server.contract;

public record NativeDocumentTaskResponse(String taskId, String state, long uploadedBytes, long byteLength,
        WorkbookImportResponse result, String errorCode, String errorMessage) { }
