package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonValue;

import java.util.Locale;

public enum WorkbookAclRole {
    OWNER(4),
    EDITOR(3),
    COMMENTER(2),
    VIEWER(1);

    private final int rank;

    WorkbookAclRole(int rank) {
        this.rank = rank;
    }

    public boolean includes(WorkbookAclRole required) {
        return rank >= required.rank;
    }

    /**
     * Role labels are a wire contract, not Java enum names.  The browser
     * contract uses lower-case labels while the database keeps enum names for
     * its check constraint.
     */
    @JsonValue
    public String wireValue() {
        return name().toLowerCase(Locale.ROOT);
    }

    @JsonCreator
    public static WorkbookAclRole fromWireValue(String value) {
        if (value == null || value.isBlank()) throw new IllegalArgumentException("Workbook role is required");
        try {
            return valueOf(value.trim().toUpperCase(Locale.ROOT));
        } catch (IllegalArgumentException error) {
            throw new IllegalArgumentException("Unknown workbook role: " + value);
        }
    }
}
