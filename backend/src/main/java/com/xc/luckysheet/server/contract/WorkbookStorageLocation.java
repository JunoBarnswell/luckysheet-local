package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

import java.util.Locale;

public enum WorkbookStorageLocation {
    LOCAL,
    REMOTE,
    MIRRORED;

    @JsonValue
    public String wireValue() {
        return name().toLowerCase(Locale.ROOT);
    }

    @JsonCreator
    public static WorkbookStorageLocation fromWireValue(String value) {
        if (value == null || value.isBlank()) throw new IllegalArgumentException("Workbook storage location is required");
        try { return valueOf(value.trim().toUpperCase(Locale.ROOT)); }
        catch (IllegalArgumentException error) { throw new IllegalArgumentException("Unknown workbook storage location: " + value); }
    }
}
