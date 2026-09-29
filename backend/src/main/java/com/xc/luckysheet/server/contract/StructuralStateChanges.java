package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonSetter;
import com.fasterxml.jackson.annotation.Nulls;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.NullNode;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Exact snapshot facts produced by the Java structural planner. No coordinate
 * transform, formula rewrite, or inverse intent is required to apply them.
 * Absent properties and explicit JSON null are deliberately distinct. These
 * are server-authored output facts, never a client-authorized write request.
 */
@JsonInclude(JsonInclude.Include.ALWAYS)
public record StructuralStateChanges(
        @JsonProperty(value = "unitId", required = true) String unitId,
        @JsonProperty(value = "sheetOrderBefore", required = true) List<String> sheetOrderBefore,
        @JsonProperty(value = "sheetOrderAfter", required = true) List<String> sheetOrderAfter,
        @JsonProperty(value = "workbookProperties", required = true) List<PropertyChange> workbookProperties,
        @JsonProperty(value = "sheets", required = true) List<SheetChange> sheets
) {
    public StructuralStateChanges {
        requireId(unitId, "workbook");
        sheetOrderBefore = order(sheetOrderBefore);
        sheetOrderAfter = order(sheetOrderAfter);
        workbookProperties = properties(workbookProperties, Set.of("schema", "version", "unitId", "sheets"));
        if (sheets == null) throw new IllegalArgumentException("Structural worksheet changes are required");
        sheets = List.copyOf(sheets);
        Set<String> beforeIds = new HashSet<>(sheetOrderBefore);
        Set<String> afterIds = new HashSet<>(sheetOrderAfter);
        Set<String> changed = new HashSet<>();
        for (SheetChange sheet : sheets) {
            if (!changed.add(sheet.sheetId())
                    || sheet.beforePresent() != beforeIds.contains(sheet.sheetId())
                    || sheet.afterPresent() != afterIds.contains(sheet.sheetId())) {
                throw new IllegalArgumentException("Structural sheet changes disagree with sheet order");
            }
        }
        Set<String> lifecycle = new HashSet<>(sheetOrderBefore);
        for (String sheetId : sheetOrderAfter) {
            if (!lifecycle.remove(sheetId)) lifecycle.add(sheetId);
        }
        if (!changed.containsAll(lifecycle)) {
            throw new IllegalArgumentException("Structural sheet lifecycle facts are missing");
        }
    }

    public StructuralStateChanges inverse() {
        return new StructuralStateChanges(unitId, sheetOrderAfter, sheetOrderBefore,
                workbookProperties.stream().map(PropertyChange::inverse).toList(),
                sheets.stream().map(SheetChange::inverse).toList());
    }

    private static List<String> order(List<String> input) {
        if (input == null) throw new IllegalArgumentException("Structural worksheet order is required");
        List<String> result = List.copyOf(input);
        if (result.isEmpty()) throw new IllegalArgumentException("A workbook must retain at least one worksheet");
        Set<String> seen = new HashSet<>();
        for (String id : result) {
            requireId(id, "sheet");
            if (!seen.add(id)) throw new IllegalArgumentException("Structural sheet order contains duplicate ids");
        }
        return result;
    }

    private static void requireId(String value, String label) {
        if (value == null || value.isBlank()) throw new IllegalArgumentException("Structural " + label + " id is required");
    }

    private static List<PropertyChange> properties(List<PropertyChange> input, Set<String> reserved) {
        if (input == null) throw new IllegalArgumentException("Structural property changes are required");
        List<PropertyChange> result = List.copyOf(input);
        Set<String> names = new HashSet<>();
        for (PropertyChange property : result) {
            if (reserved.contains(property.name()) || !names.add(property.name())) {
                throw new IllegalArgumentException("Structural property is reserved or repeated: " + property.name());
            }
        }
        return result;
    }

    @JsonInclude(JsonInclude.Include.ALWAYS)
    public record PropertyValue(
            @JsonProperty(value = "present", required = true) Boolean present,
            @JsonProperty(value = "value", required = true) JsonNode value
    ) {
        public PropertyValue {
            if (present == null) throw new IllegalArgumentException("Structural property presence is required");
            value = value == null ? NullNode.instance : value.deepCopy();
            if (!present && !value.isNull()) throw new IllegalArgumentException("An absent structural property cannot carry a value");
        }

        public static PropertyValue of(JsonNode value) {
            return new PropertyValue(value != null, value);
        }
    }

    public record PropertyChange(
            @JsonProperty(value = "name", required = true) String name,
            @JsonProperty(value = "before", required = true) PropertyValue before,
            @JsonProperty(value = "after", required = true) PropertyValue after
    ) {
        public PropertyChange {
            requireId(name, "property");
            if (before == null || after == null || before.equals(after)) {
                throw new IllegalArgumentException("Structural property requires distinct before and after facts");
            }
        }

        public PropertyChange inverse() {
            return new PropertyChange(name, after, before);
        }
    }

    public record SheetChange(
            @JsonProperty(value = "sheetId", required = true) String sheetId,
            @JsonProperty(value = "beforePresent", required = true) Boolean beforePresent,
            @JsonProperty(value = "afterPresent", required = true) Boolean afterPresent,
            @JsonProperty(value = "properties", required = true) List<PropertyChange> properties,
            @JsonProperty(value = "rows", required = true) List<CellRowChange> rows
    ) {
        public SheetChange {
            requireId(sheetId, "sheet");
            if (beforePresent == null || afterPresent == null || rows == null) {
                throw new IllegalArgumentException("Structural worksheet presence and rows are required");
            }
            if (!beforePresent && !afterPresent) throw new IllegalArgumentException("Structural sheet has no state");
            properties = StructuralStateChanges.properties(properties, Set.of("cells"));
            rows = List.copyOf(rows);
            Set<Integer> rowIds = new HashSet<>();
            for (CellRowChange row : rows) {
                if (!rowIds.add(row.row()) || (!beforePresent && row.beforePresent()) || (!afterPresent && row.afterPresent())) {
                    throw new IllegalArgumentException("Structural cell row disagrees with worksheet lifecycle");
                }
            }
            for (PropertyChange property : properties) {
                if ((!beforePresent && property.before().present()) || (!afterPresent && property.after().present())) {
                    throw new IllegalArgumentException("Structural property disagrees with worksheet lifecycle");
                }
                if ("id".equals(property.name()) && (beforePresent.equals(afterPresent)
                        || (beforePresent && !sheetId.equals(property.before().value().textValue()))
                        || (afterPresent && !sheetId.equals(property.after().value().textValue())))) {
                    throw new IllegalArgumentException("Worksheet identity cannot be rewritten");
                }
            }
            if (!beforePresent.equals(afterPresent) && properties.stream().noneMatch(property -> "id".equals(property.name()))) {
                throw new IllegalArgumentException("Worksheet lifecycle requires an identity fact");
            }
            if (beforePresent.equals(afterPresent) && properties.isEmpty() && rows.isEmpty()) {
                throw new IllegalArgumentException("Empty structural worksheet change");
            }
        }

        public SheetChange inverse() {
            return new SheetChange(sheetId, afterPresent, beforePresent,
                    properties.stream().map(PropertyChange::inverse).toList(),
                    rows.stream().map(CellRowChange::inverse).toList());
        }
    }

    /** Row presence preserves even an explicitly empty stored row without scanning its address space. */
    public record CellRowChange(
            @JsonProperty(value = "row", required = true) @JsonSetter(nulls = Nulls.FAIL) int row,
            @JsonProperty(value = "beforePresent", required = true) Boolean beforePresent,
            @JsonProperty(value = "afterPresent", required = true) Boolean afterPresent,
            @JsonProperty(value = "cells", required = true) List<CellChange> cells
    ) {
        public CellRowChange {
            if (beforePresent == null || afterPresent == null || cells == null) {
                throw new IllegalArgumentException("Structural row presence and cells are required");
            }
            if (row < 0 || row > 1_048_575 || (!beforePresent && !afterPresent)) {
                throw new IllegalArgumentException("Structural cell row is invalid");
            }
            cells = List.copyOf(cells);
            Set<Integer> columns = new HashSet<>();
            for (CellChange cell : cells) {
                if (!columns.add(cell.column()) || (!beforePresent && cell.before() != null)
                        || (!afterPresent && cell.after() != null)) {
                    throw new IllegalArgumentException("Structural cell disagrees with its row lifecycle");
                }
            }
            if (beforePresent.equals(afterPresent) && cells.isEmpty()) throw new IllegalArgumentException("Empty structural cell row change");
        }

        public CellRowChange inverse() {
            return new CellRowChange(row, afterPresent, beforePresent, cells.stream().map(CellChange::inverse).toList());
        }
    }

    @JsonInclude(JsonInclude.Include.ALWAYS)
    public record CellChange(
            @JsonProperty(value = "column", required = true) @JsonSetter(nulls = Nulls.FAIL) int column,
            @JsonProperty(value = "before", required = true) JsonNode before,
            @JsonProperty(value = "after", required = true) JsonNode after
    ) {
        public CellChange {
            // Jackson represents an absent cell as JSON null on the wire.
            before = before == null || before.isNull() ? null : before.deepCopy();
            after = after == null || after.isNull() ? null : after.deepCopy();
            if (column < 0 || column > 16_383 || (before == null && after == null)
                    || (before != null && !before.isObject()) || (after != null && !after.isObject())
                    || java.util.Objects.equals(before, after)) {
                throw new IllegalArgumentException("Structural cell requires distinct object states within worksheet bounds");
            }
        }

        public CellChange inverse() {
            return new CellChange(column, after, before);
        }
    }
}
