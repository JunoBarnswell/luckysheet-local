package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

import java.util.Locale;

public enum WorkbookSyncStatus {
    SYNCED,
    SYNCING,
    PENDING,
    OFFLINE,
    CONFLICT,
    ERROR;

    @JsonValue
    public String wireValue() {
        return name().toLowerCase(Locale.ROOT);
    }

    @JsonCreator
    public static WorkbookSyncStatus fromWireValue(String value) {
        if (value == null || value.isBlank()) throw new IllegalArgumentException("Workbook sync status is required");
        try { return valueOf(value.trim().toUpperCase(Locale.ROOT)); }
        catch (IllegalArgumentException error) { throw new IllegalArgumentException("Unknown workbook sync status: " + value); }
    }
}
