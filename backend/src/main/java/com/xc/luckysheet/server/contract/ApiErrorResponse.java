package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;

public record ApiErrorResponse(String code, String message, JsonNode details) {
    public ApiErrorResponse(String code, String message) {
        this(code, message, null);
    }
}
