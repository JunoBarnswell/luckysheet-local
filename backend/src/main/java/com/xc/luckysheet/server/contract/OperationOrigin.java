package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonValue;

public enum OperationOrigin {
    CLIENT,
    SYSTEM;

    @JsonValue
    public String wireValue() {
        return name().toLowerCase(java.util.Locale.ROOT);
    }
}
