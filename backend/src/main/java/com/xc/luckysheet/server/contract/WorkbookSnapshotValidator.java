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
        if (snapshot.path("version").asInt(-1) != 2 || !snapshot.path("sheets").isArray()) {
            throw ServiceException.validation("Stored workbook snapshot version is invalid");
        }
        snapshot.put("version", GeneratedWorkbookContract.SNAPSHOT_VERSION);
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
