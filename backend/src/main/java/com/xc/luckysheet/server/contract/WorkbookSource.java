package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

import java.util.Locale;

public enum WorkbookSource {
    NATIVE,
    DOCUMENT_IMPORT;

    @JsonValue
    public String wireValue() {
        return name().toLowerCase(Locale.ROOT).replace('_', '-');
    }

    @JsonCreator
    public static WorkbookSource fromWireValue(String value) {
        if (value == null || value.isBlank()) throw new IllegalArgumentException("Workbook source is required");
        return switch (value.trim().toLowerCase(Locale.ROOT)) {
            case "native" -> NATIVE;
            case "document-import", "document_import" -> DOCUMENT_IMPORT;
            default -> throw new IllegalArgumentException("Unknown workbook source: " + value);
        };
    }
}
