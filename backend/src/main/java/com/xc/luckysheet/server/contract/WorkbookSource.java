package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

import java.util.Locale;

public enum WorkbookSource {
    NATIVE,
    XLSX_IMPORT;

    @JsonValue
    public String wireValue() {
        return name().toLowerCase(Locale.ROOT).replace('_', '-');
    }

    @JsonCreator
    public static WorkbookSource fromWireValue(String value) {
        if (value == null || value.isBlank()) throw new IllegalArgumentException("Workbook source is required");
        return switch (value.trim().toLowerCase(Locale.ROOT)) {
            case "native" -> NATIVE;
            case "xlsx-import", "xlsx_import" -> XLSX_IMPORT;
            default -> throw new IllegalArgumentException("Unknown workbook source: " + value);
        };
    }
}
