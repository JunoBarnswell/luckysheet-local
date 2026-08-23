package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

import java.util.Locale;

public enum WorkbookLifecycle {
    ACTIVE,
    TRASHED;

    @JsonValue
    public String wireValue() {
        return name().toLowerCase(Locale.ROOT);
    }

    @JsonCreator
    public static WorkbookLifecycle fromWireValue(String value) {
        if (value == null || value.isBlank()) throw new IllegalArgumentException("Workbook lifecycle is required");
        try { return valueOf(value.trim().toUpperCase(Locale.ROOT)); }
        catch (IllegalArgumentException error) { throw new IllegalArgumentException("Unknown workbook lifecycle: " + value); }
    }
}
