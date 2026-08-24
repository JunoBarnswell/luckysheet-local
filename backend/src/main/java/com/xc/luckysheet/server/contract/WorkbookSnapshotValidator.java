package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.xc.luckysheet.server.service.ServiceException;

/**
 * Server-side wire validator for the canonical WorkbookSnapshot. The browser
 * has a richer type model, but persistence must reject malformed snapshots
 * before they become a historical checkpoint.
 */
public final class WorkbookSnapshotValidator {
    private WorkbookSnapshotValidator() {
    }

    public static ObjectNode requireCanonical(JsonNode value, String expectedUnitId) {
        if (value == null || !value.isObject()) throw ServiceException.validation("Workbook snapshot must be an object");
        ObjectNode snapshot = (ObjectNode) value;
        if (!GeneratedWorkbookContract.SNAPSHOT_SCHEMA.equals(snapshot.path("schema").asText())) {
            throw ServiceException.validation("Workbook snapshot schema is invalid");
        }
        if (!snapshot.path("version").canConvertToInt()
                || snapshot.path("version").intValue() != GeneratedWorkbookContract.SNAPSHOT_VERSION) {
            throw ServiceException.validation("Workbook snapshot version is invalid");
        }
        String unitId = snapshot.path("unitId").asText().trim();
        if (unitId.isBlank() || !unitId.equals(expectedUnitId)) {
            throw ServiceException.validation("Workbook snapshot unitId does not match the request");
        }
        String name = snapshot.path("name").asText().trim();
        if (name.isBlank() || name.length() > GeneratedWorkbookContract.MAX_WORKBOOK_NAME_LENGTH) {
            throw ServiceException.validation("Workbook snapshot name is invalid");
        }
        JsonNode dimensionMetrics = snapshot.get("dimensionMetrics");
        if (dimensionMetrics == null || !dimensionMetrics.isObject() || dimensionMetrics.path("normalFontFamily").asText().isBlank()
                || !dimensionMetrics.path("normalFontSizePx").isNumber() || dimensionMetrics.path("normalFontSizePx").asDouble() <= 0
                || !dimensionMetrics.path("maximumDigitWidthPx").isNumber() || dimensionMetrics.path("maximumDigitWidthPx").asDouble() <= 0) {
            throw ServiceException.validation("Workbook snapshot dimensionMetrics is invalid");
        }
        JsonNode sheets = snapshot.get("sheets");
        if (sheets == null || !sheets.isArray() || sheets.isEmpty()) {
            throw ServiceException.validation("Workbook snapshot requires at least one sheet");
        }
        if (!snapshot.path("dataSources").isArray()) {
            throw ServiceException.validation("Workbook snapshot dataSources must be an array");
        }
        java.util.Set<String> sheetIds = new java.util.HashSet<>();
        for (JsonNode sheet : sheets) {
            if (!sheet.isObject()) throw ServiceException.validation("Workbook snapshot sheet is invalid");
            String sheetId = sheet.path("id").asText().trim();
            String sheetName = sheet.path("name").asText().trim();
            if (sheetId.isBlank() || sheetName.isBlank() || !sheetIds.add(sheetId)) {
                throw ServiceException.validation("Workbook snapshot sheet identity is invalid");
            }
            if (!sheet.path("rowCount").canConvertToInt() || sheet.path("rowCount").intValue() < 1
                    || !sheet.path("columnCount").canConvertToInt() || sheet.path("columnCount").intValue() < 1
                    || !sheet.path("cells").isObject() || !sheet.path("merges").isArray()
                    || !sheet.path("pivots").isArray() || !sheet.path("sparklines").isArray()
                    || !sheet.path("drawings").isArray() || !sheet.path("drawingPayloads").isObject()) {
                throw ServiceException.validation("Workbook snapshot sheet grid is invalid");
            }
            if (!sheet.path("defaultRowHeightPx").isNumber() || sheet.path("defaultRowHeightPx").asDouble() <= 0
                    || !sheet.path("defaultColumnWidthPx").isNumber() || sheet.path("defaultColumnWidthPx").asDouble() <= 0
                    || !sheet.path("pane").isObject()
                    || !("none".equals(sheet.path("pane").path("kind").asText())
                    || "frozen".equals(sheet.path("pane").path("kind").asText())
                    || "split".equals(sheet.path("pane").path("kind").asText()))) {
                throw ServiceException.validation("Workbook snapshot sheet pixel geometry is invalid");
            }
            JsonNode pane = sheet.path("pane");
            if (!"none".equals(pane.path("kind").asText())) {
                String state = pane.path("state").asText();
                if (("frozen".equals(pane.path("kind").asText()) && !("frozen".equals(state) || "frozenSplit".equals(state)))
                        || ("split".equals(pane.path("kind").asText()) && !"split".equals(state))) {
                    throw ServiceException.validation("Workbook snapshot pane state is invalid");
                }
            }
            JsonNode autoFilter = sheet.get("autoFilter");
            RangeRef worksheetFilterRange = autoFilter == null || autoFilter.isNull() ? null : validateAutoFilter(autoFilter, sheetId, null);
            java.util.List<RangeRef> tableFilterRanges = new java.util.ArrayList<>();
            JsonNode tables = sheet.get("sheetTables");
            if (tables != null && !tables.isNull()) {
                if (!tables.isArray()) throw ServiceException.validation("Workbook snapshot sheetTables is invalid");
                for (JsonNode table : tables) {
                    if (!table.isObject()) throw ServiceException.validation("Workbook snapshot table is invalid");
                    JsonNode tableFilter = table.get("autoFilter");
                    if (tableFilter == null || tableFilter.isNull()) continue;
                    RangeRef tableRange = rangeOf(table.get("range"), sheetId);
                    RangeRef filterRange = validateAutoFilter(tableFilter, sheetId, tableRange);
                    if (worksheetFilterRange != null && overlaps(worksheetFilterRange, filterRange)) {
                        throw ServiceException.validation("Worksheet and Table AutoFilter ranges cannot overlap");
                    }
                    if (tableFilterRanges.stream().anyMatch(existing -> overlaps(existing, filterRange))) {
                        throw ServiceException.validation("Table AutoFilter ranges cannot overlap");
                    }
                    tableFilterRanges.add(filterRange);
                }
            }
        }
        return snapshot;
    }

    /** One-way migration used only when reading persisted v2 checkpoints. */
    public static ObjectNode migrateStored(JsonNode value, String expectedUnitId) {
        if (value == null || !value.isObject()) throw ServiceException.validation("Stored workbook snapshot must be an object");
        ObjectNode snapshot = ((ObjectNode) value).deepCopy();
        if (snapshot.path("version").asInt(-1) == GeneratedWorkbookContract.SNAPSHOT_VERSION) {
            return requireCanonical(snapshot, expectedUnitId);
        }
        if (snapshot.path("version").asInt(-1) == 3 && snapshot.path("sheets").isArray()) {
            snapshot.put("version", GeneratedWorkbookContract.SNAPSHOT_VERSION);
            for (JsonNode raw : (ArrayNode) snapshot.path("sheets")) {
                if (!raw.isObject()) throw ServiceException.validation("Stored workbook snapshot sheet is invalid");
                ObjectNode sheet = (ObjectNode) raw;
                JsonNode pane = sheet.get("pane");
                if (pane != null && pane.isObject() && !"none".equals(pane.path("kind").asText())) {
                    ObjectNode paneObject = (ObjectNode) pane;
                    String kind = pane.path("kind").asText();
                    if (!paneObject.has("state")) paneObject.put("state", "split".equals(kind) ? "split" : "frozen");
                    if (!paneObject.has("startRow")) paneObject.put("startRow", pane.path("ySplit").asInt(0));
                    if (!paneObject.has("startColumn")) paneObject.put("startColumn", pane.path("xSplit").asInt(0));
                }
                JsonNode legacy = sheet.get("filter");
                if (legacy == null || !legacy.isObject()) continue;
                ObjectNode autoFilter = snapshot.objectNode();
                autoFilter.put("sheetId", legacy.path("sheetId").asText(sheet.path("id").asText()));
                autoFilter.set("range", legacy.path("range").deepCopy());
                ObjectNode columns = snapshot.objectNode();
                legacy.path("criteria").fields().forEachRemaining(entry -> {
                    int columnIndex = Integer.parseInt(entry.getKey());
                    ObjectNode column = snapshot.objectNode().put("column", columnIndex).put("showButton", true).put("hiddenButton", false);
                    JsonNode selected = entry.getValue().get("selectedValues");
                    if (selected != null && selected.isArray()) {
                        boolean includeBlank = false;
                        for (JsonNode selectedValue : selected) includeBlank |= selectedValue.isNull() || selectedValue.asText().isEmpty();
                        ObjectNode criterion = snapshot.objectNode().put("kind", "values").put("includeBlank", includeBlank);
                        criterion.set("values", selected.deepCopy());
                        column.set("criterion", criterion);
                    } else if (entry.getValue().has("conditionOperator")) {
                        ObjectNode criterion = snapshot.objectNode().put("kind", "custom").put("join", "and");
                        ArrayNode conditions = snapshot.arrayNode();
                        conditions.add(snapshot.objectNode().put("operator", entry.getValue().path("conditionOperator").asText()).set("value", entry.getValue().get("conditionValue")));
                        criterion.set("conditions", conditions);
                        column.set("criterion", criterion);
                    }
                    columns.set(entry.getKey(), column);
                });
                autoFilter.set("columns", columns);
                sheet.set("autoFilter", autoFilter);
                sheet.remove("filter");
            }
            return requireCanonical(snapshot, expectedUnitId);
        }
        if (snapshot.path("version").asInt(-1) != 2 || !snapshot.path("sheets").isArray()) {
            throw ServiceException.validation("Stored workbook snapshot version is invalid");
        }
        snapshot.put("version", GeneratedWorkbookContract.SNAPSHOT_VERSION);
        snapshot.putObject("dimensionMetrics").put("normalFontFamily", "Calibri").put("normalFontSizePx", 14.6666666667).put("maximumDigitWidthPx", 7);
        for (JsonNode raw : (ArrayNode) snapshot.path("sheets")) {
            if (!raw.isObject()) throw ServiceException.validation("Stored workbook snapshot sheet is invalid");
            ObjectNode sheet = (ObjectNode) raw;
            sheet.put("defaultRowHeightPx", positiveOr(sheet.get("defaultRowHeight"), 28));
            sheet.put("defaultColumnWidthPx", positiveOr(sheet.get("defaultColumnWidth"), 110));
            sheet.set("rowHeightsPx", copyObjectOrEmpty(sheet.get("rowHeights")));
            sheet.set("columnWidthsPx", copyObjectOrEmpty(sheet.get("columnWidths")));
            JsonNode freeze = sheet.get("freeze");
            int xSplit = freeze == null ? 0 : freeze.path("xSplit").asInt(0);
            int ySplit = freeze == null ? 0 : freeze.path("ySplit").asInt(0);
            ObjectNode pane = sheet.putObject("pane");
            if (xSplit > 0 || ySplit > 0) {
                pane.put("kind", "frozen").put("xSplit", xSplit).put("ySplit", ySplit)
                        .put("startRow", freeze.path("startRow").asInt(ySplit))
                        .put("startColumn", freeze.path("startColumn").asInt(xSplit)).put("state", "frozen");
            } else pane.put("kind", "none");
            migrateFontSizes(sheet);
            sheet.remove(java.util.List.of("defaultRowHeight", "defaultColumnWidth", "rowHeights", "columnWidths", "freeze"));
        }
        return requireCanonical(snapshot, expectedUnitId);
    }

    private static double positiveOr(JsonNode value, double fallback) {
        return value != null && value.isNumber() && value.asDouble() > 0 ? value.asDouble() : fallback;
    }

    private static RangeRef validateAutoFilter(JsonNode filter, String sheetId, RangeRef expectedRange) {
        if (!filter.isObject() || !sheetId.equals(filter.path("sheetId").asText())) {
            throw ServiceException.validation("Workbook snapshot autoFilter is invalid");
        }
        RangeRef range = rangeOf(filter.get("range"), sheetId);
        if (expectedRange != null && !sameRange(range, expectedRange)) {
            throw ServiceException.validation("Table AutoFilter range must equal the Table range");
        }
        JsonNode columns = filter.get("columns");
        if (columns == null || !columns.isObject()) throw ServiceException.validation("AutoFilter columns must be an object");
        columns.fields().forEachRemaining(entry -> {
            try {
                int key = Integer.parseInt(entry.getKey());
                JsonNode column = entry.getValue();
                if (!column.isObject() || column.path("column").asInt(Integer.MIN_VALUE) != key
                        || key < range.startColumn() || key > range.endColumn()
                        || !column.path("showButton").isBoolean() || !column.path("hiddenButton").isBoolean()) {
                    throw ServiceException.validation("AutoFilter column identity is invalid");
                }
                JsonNode criterion = column.get("criterion");
                if (criterion != null && !criterion.isNull()
                        && (!criterion.isObject() || !java.util.Set.of("values", "custom", "dynamic", "top10", "color", "icon").contains(criterion.path("kind").asText()))) {
                    throw ServiceException.validation("AutoFilter criterion kind is invalid");
                }
            } catch (NumberFormatException error) {
                throw ServiceException.validation("AutoFilter column key is invalid");
            }
        });
        return range;
    }

    private static RangeRef rangeOf(JsonNode value, String sheetId) {
        if (value == null || !value.isObject() || !sheetId.equals(value.path("sheetId").asText())) {
            throw ServiceException.validation("AutoFilter range is invalid");
        }
        try {
            return new RangeRef(sheetId, value.path("startRow").asInt(), value.path("endRow").asInt(), value.path("startColumn").asInt(), value.path("endColumn").asInt());
        } catch (IllegalArgumentException error) {
            throw ServiceException.validation("AutoFilter range is invalid");
        }
    }

    private static boolean sameRange(RangeRef left, RangeRef right) {
        return left.sheetId().equals(right.sheetId()) && left.startRow() == right.startRow() && left.endRow() == right.endRow()
                && left.startColumn() == right.startColumn() && left.endColumn() == right.endColumn();
    }

    private static boolean overlaps(RangeRef left, RangeRef right) {
        return left.startRow() <= right.endRow() && right.startRow() <= left.endRow()
                && left.startColumn() <= right.endColumn() && right.startColumn() <= left.endColumn();
    }

    private static ObjectNode copyObjectOrEmpty(JsonNode value) {
        return value != null && value.isObject() ? ((ObjectNode) value).deepCopy() : com.fasterxml.jackson.databind.node.JsonNodeFactory.instance.objectNode();
    }

    private static void migrateFontSizes(JsonNode value) {
        if (value == null) return;
        if (value.isArray()) value.forEach(WorkbookSnapshotValidator::migrateFontSizes);
        if (!value.isObject()) return;
        ObjectNode object = (ObjectNode) value;
        if (object.path("fontSize").isNumber() && !object.has("fontSizePx")) object.set("fontSizePx", object.get("fontSize"));
        object.remove("fontSize");
        java.util.List<JsonNode> children = new java.util.ArrayList<>();
        object.elements().forEachRemaining(children::add);
        children.forEach(WorkbookSnapshotValidator::migrateFontSizes);
    }
}
