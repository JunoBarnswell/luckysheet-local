package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

import java.util.Locale;

public enum WorkspaceSpaceType {
    PERSONAL,
    TEAM;

    @JsonValue
    public String wireValue() {
        return name().toLowerCase(Locale.ROOT);
    }

    @JsonCreator
    public static WorkspaceSpaceType fromWireValue(String value) {
        if (value == null || value.isBlank()) throw new IllegalArgumentException("Space type is required");
        try {
            return valueOf(value.trim().toUpperCase(Locale.ROOT));
        } catch (IllegalArgumentException error) {
            throw new IllegalArgumentException("Unknown space type: " + value);
        }
    }
}
