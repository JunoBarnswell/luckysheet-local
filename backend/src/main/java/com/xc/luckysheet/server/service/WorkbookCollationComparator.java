package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;

import java.math.BigDecimal;

/** One deterministic value ordering for server-side workbook/query sorting. */
final class WorkbookCollationComparator {
    private WorkbookCollationComparator() {
    }

    static int compare(JsonNode left, JsonNode right) {
        Kind leftKind = kind(left);
        Kind rightKind = kind(right);
        if (leftKind != rightKind) return Integer.compare(leftKind.rank, rightKind.rank);
        return switch (leftKind) {
            case NUMBER -> number(left, right);
            case TEXT -> left.asText().compareTo(right.asText());
            case BOOLEAN -> Boolean.compare(left.asBoolean(), right.asBoolean());
            case ERROR -> left.asText().compareTo(right.asText());
            case BLANK -> 0;
        };
    }

    private static int number(JsonNode left, JsonNode right) {
        try {
            return new BigDecimal(left.asText()).compareTo(new BigDecimal(right.asText()));
        } catch (NumberFormatException ignored) {
            return Double.compare(left.asDouble(), right.asDouble());
        }
    }

    private static Kind kind(JsonNode value) {
        if (value == null || value.isNull()) return Kind.BLANK;
        if (value.isNumber()) return Kind.NUMBER;
        if (value.isBoolean()) return Kind.BOOLEAN;
        if (value.isTextual()) return value.asText().matches("#(NULL!|DIV/0!|VALUE!|REF!|NAME\\?|NUM!|N/A|CALC!|BLOCKED!|SPILL!|PARSE!|CYCLE!)") ? Kind.ERROR : Kind.TEXT;
        return Kind.TEXT;
    }

    private enum Kind {
        NUMBER(0),
        TEXT(1),
        BOOLEAN(2),
        ERROR(3),
        BLANK(4);

        private final int rank;

        Kind(int rank) {
            this.rank = rank;
        }
    }
}
