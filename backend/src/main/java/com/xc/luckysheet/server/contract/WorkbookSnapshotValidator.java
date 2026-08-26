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
        JsonNode dataModel = snapshot.get("dataModel");
        if (dataModel == null || !dataModel.isObject() || !dataModel.path("sources").isArray()
                || !dataModel.path("tables").isArray() || !dataModel.path("relationships").isArray() || !dataModel.path("views").isArray()) {
            throw ServiceException.validation("Workbook snapshot dataModel is invalid");
        }
        java.util.Set<String> sheetIds = new java.util.HashSet<>();
        java.util.Map<String, int[]> sheetDimensions = new java.util.HashMap<>();
        for (JsonNode sheet : sheets) {
            String sheetId = sheet.path("id").asText().trim();
            if (!sheetId.isBlank() && sheet.path("rowCount").canConvertToInt() && sheet.path("columnCount").canConvertToInt()) {
                sheetDimensions.put(sheetId, new int[]{sheet.path("rowCount").intValue(), sheet.path("columnCount").intValue()});
            }
        }
        java.util.Set<String> pivotIds = new java.util.HashSet<>();
        java.util.Map<String, JsonNode> pivotsById = new java.util.HashMap<>();
        java.util.Map<String, String> pivotSourceKeys = new java.util.HashMap<>();
        for (JsonNode sheet : sheets) {
            if (!sheet.isObject()) throw ServiceException.validation("Workbook snapshot sheet is invalid");
            String sheetId = sheet.path("id").asText().trim();
            String sheetName = sheet.path("name").asText().trim();
            if (sheetId.isBlank() || sheetName.isBlank() || !sheetIds.add(sheetId)) {
                throw ServiceException.validation("Workbook snapshot sheet identity is invalid");
            }
            String sheetKind = sheet.path("kind").asText();
            if (!java.util.Set.of("worksheet", "table-sheet", "gantt-sheet", "report-sheet").contains(sheetKind)) throw ServiceException.validation("Workbook snapshot sheet kind is invalid");
            if ("table-sheet".equals(sheetKind) && !sheet.path("tableSheet").isObject()) throw ServiceException.validation("TableSheet definition is required");
            if ("gantt-sheet".equals(sheetKind) && !sheet.path("ganttSheet").isObject()) throw ServiceException.validation("GanttSheet definition is required");
            if ("report-sheet".equals(sheetKind) && !sheet.path("reportSheet").isObject()) throw ServiceException.validation("ReportSheet definition is required");
            if (!sheet.path("rowCount").canConvertToInt() || sheet.path("rowCount").intValue() < 1
                    || !sheet.path("columnCount").canConvertToInt() || sheet.path("columnCount").intValue() < 1
                    || !sheet.path("cells").isObject() || !sheet.path("merges").isArray()
                    || !sheet.path("pivots").isArray() || !sheet.path("sparklines").isArray()
                    || !sheet.path("drawings").isArray() || !sheet.path("drawingPayloads").isObject()) {
                throw ServiceException.validation("Workbook snapshot sheet grid is invalid");
            }
            validateReviewSnapshot(sheet.get("review"), sheetId);
            for (JsonNode pivot : sheet.path("pivots")) {
                if (!pivot.isObject() || pivot.path("id").asText().isBlank() || !pivotIds.add(pivot.path("id").asText())) {
                    throw ServiceException.validation("Workbook snapshot Pivot identity is duplicated or empty");
                }
                pivotsById.put(pivot.path("id").asText(), pivot);
            }
            if (!sheet.path("defaultRowHeightPx").isNumber() || sheet.path("defaultRowHeightPx").asDouble() <= 0
                    || !sheet.path("defaultColumnWidthPx").isNumber() || sheet.path("defaultColumnWidthPx").asDouble() <= 0
                    || !sheet.path("pane").isObject()
                    || !("none".equals(sheet.path("pane").path("kind").asText())
                    || "frozen".equals(sheet.path("pane").path("kind").asText())
                    || "split".equals(sheet.path("pane").path("kind").asText()))) {
                throw ServiceException.validation("Workbook snapshot sheet pixel geometry is invalid");
            }
            sheet.path("drawingPayloads").fields().forEachRemaining(entry -> {
                JsonNode payload = entry.getValue();
                if ("camera".equals(payload.path("kind").asText())) {
                    validateDrawingSourceRange(payload.get("sourceRange"), sheetDimensions, "Camera");
                }
                if ("image".equals(payload.path("kind").asText())) validateAssetRef(payload.get("asset"), "Drawing image");
            });
            sheet.path("cells").fields().forEachRemaining(row -> row.getValue().fields().forEachRemaining(cell -> {
                JsonNode presentation = cell.getValue().get("presentation");
                if (presentation != null && "image".equals(presentation.path("kind").asText())) validateAssetRef(presentation.get("asset"), "Cell image");
            }));
            JsonNode pane = sheet.path("pane");
            if (!"none".equals(pane.path("kind").asText())) {
                String state = pane.path("state").asText();
                if (("frozen".equals(pane.path("kind").asText()) && !("frozen".equals(state) || "frozenSplit".equals(state)))
                        || ("split".equals(pane.path("kind").asText()) && !"split".equals(state))) {
                    throw ServiceException.validation("Workbook snapshot pane state is invalid");
                }
            }
            JsonNode autoFilter = sheet.get("autoFilter");
            if (autoFilter != null && !autoFilter.isNull()) validateAutoFilter(autoFilter, sheetId, null);
            JsonNode tables = sheet.get("sheetTables");
            if (tables != null && !tables.isNull()) {
                if (!tables.isArray()) throw ServiceException.validation("Workbook snapshot sheetTables is invalid");
                for (JsonNode table : tables) {
                    if (!table.isObject()) throw ServiceException.validation("Workbook snapshot table is invalid");
                    JsonNode tableFilter = table.get("autoFilter");
                    if (tableFilter == null || tableFilter.isNull()) continue;
                    RangeRef tableRange = rangeOf(table.get("range"), sheetId);
                    validateAutoFilter(tableFilter, sheetId, tableRange);
                }
            }
            AutoFilterOwnershipValidator.resolveOwners((ObjectNode) sheet, sheetId);
        }
        for (JsonNode sheet : sheets) {
            ObjectNode sheetObject = (ObjectNode) sheet;
            ObjectNode payloads = (ObjectNode) sheetObject.path("drawingPayloads");
            for (JsonNode drawing : sheetObject.path("drawings")) {
                if (!drawing.isObject()) throw ServiceException.validation("Workbook snapshot drawing is invalid");
                String drawingId = drawing.path("id").asText();
                String payloadId = drawing.path("payloadId").asText();
                JsonNode payload = payloads.get(payloadId);
                if (payload == null || !payload.isObject()) throw ServiceException.validation("Drawing payload is missing: " + payloadId);
                String kind = payload.path("kind").asText();
                if (!("chart".equals(kind) || "slicer".equals(kind) || "timeline".equals(kind))) continue;
                if ("chart".equals(kind) && !payload.has("pivotId")) continue;
                String pivotId = payload.path("pivotId").asText();
                if (pivotId.isBlank() || !pivotIds.contains(pivotId)) {
                    throw ServiceException.validation("Drawing " + drawingId + " references missing Pivot: " + pivotId);
                }
                if (("slicer".equals(kind) || "timeline".equals(kind)) && payload.has("connections")) {
                    JsonNode connections = payload.get("connections");
                    if (!connections.isArray()) throw ServiceException.validation("Drawing connections is invalid: " + drawingId);
                    JsonNode primary = pivotsById.get(pivotId);
                    JsonNode primaryField = findPivotField(primary, payload.path("fieldId").asText());
                    if (primaryField == null) throw ServiceException.validation("Drawing primary field is missing: " + drawingId);
                    if (connections.isEmpty()) continue;
                    String primarySourceKey = pivotSourceKeys.computeIfAbsent(pivotId, ignored -> canonicalJson(primary.get("source")));
                    java.util.Set<String> seenConnections = new java.util.HashSet<>();
                    for (JsonNode connection : connections) {
                        if (!connection.isObject()) throw ServiceException.validation("Drawing connection is invalid: " + drawingId);
                        String targetId = connection.path("pivotId").asText();
                        String sourceKey = connection.path("sourceKey").asText();
                        String fieldId = connection.path("fieldId").asText();
                        if (targetId.isBlank() || !pivotIds.contains(targetId) || targetId.equals(pivotId) || !seenConnections.add(targetId)) {
                            throw ServiceException.validation("Drawing connection Pivot is invalid: " + drawingId);
                        }
                        JsonNode target = pivotsById.get(targetId);
                        JsonNode targetField = findPivotField(target, fieldId);
                        String targetSourceKey = pivotSourceKeys.computeIfAbsent(targetId, ignored -> canonicalJson(target.get("source")));
                        if (!primarySourceKey.equals(sourceKey) || !primarySourceKey.equals(targetSourceKey) || !compatiblePivotField(primaryField, targetField)) {
                            throw ServiceException.validation("Drawing connection source/cache/field is incompatible: " + drawingId);
                        }
                        if ("timeline".equals(kind) && (!"date".equals(primaryField.path("dataType").asText()) || !"date".equals(targetField.path("dataType").asText()))) {
                            throw ServiceException.validation("Drawing Timeline connection field is not date-semantic: " + drawingId);
                        }
                    }
                }
            }
        }
        java.util.Set<String> dataChartTableIds = new java.util.HashSet<>();
        for (JsonNode sheet : sheets) sheet.path("drawingPayloads").forEach(payload -> {
            if ("data-chart".equals(payload.path("kind").asText()) && "table".equals(payload.path("source").path("kind").asText())) {
                dataChartTableIds.add(payload.path("source").path("tableId").asText());
            }
            if ("data-chart".equals(payload.path("kind").asText()) && "report-sheet".equals(payload.path("source").path("kind").asText())) {
                validateDrawingSourceRange(payload.path("source").get("range"), sheetDimensions, "Data chart");
            }
        });
        for (JsonNode table : dataModel.path("tables")) {
            JsonNode sourceRange = table.get("sourceRange");
            if (dataChartTableIds.contains(table.path("id").asText()) && sourceRange != null && !sourceRange.isNull()) {
                validateDrawingSourceRange(sourceRange, sheetDimensions, "Data chart table");
            }
        }
        return snapshot;
    }

    private static JsonNode findPivotField(JsonNode pivot, String fieldId) {
        if (pivot == null || !pivot.path("fieldCatalog").path("fields").isArray()) return null;
        for (JsonNode field : pivot.path("fieldCatalog").path("fields")) if (field.path("fieldId").asText().equals(fieldId)) return field;
        return null;
    }

    private static void validateAssetRef(JsonNode asset, String label) {
        if (asset == null || !asset.isObject() || !"AssetRef".equals(asset.path("schema").asText())
                || asset.path("assetId").asText().isBlank() || !asset.path("contentHash").asText().matches("[a-f0-9]{64}")
                || !asset.path("mimeType").asText().startsWith("image/") || !asset.path("byteLength").canConvertToInt()
                || asset.path("byteLength").asInt(-1) < 0) {
            throw ServiceException.validation(label + " asset is invalid");
        }
    }

    private static boolean compatiblePivotField(JsonNode primary, JsonNode target) {
        return target != null
                && primary.path("ordinal").asInt(-1) == target.path("ordinal").asInt(-1)
                && primary.path("name").asText().equals(target.path("name").asText())
                && primary.path("dataType").asText().equals(target.path("dataType").asText());
    }

    private static String canonicalJson(JsonNode node) {
        if (node == null || node.isNull()) return "null";
        if (node.isObject()) {
            java.util.List<String> keys = new java.util.ArrayList<>();
            node.fieldNames().forEachRemaining(keys::add);
            java.util.Collections.sort(keys);
            StringBuilder result = new StringBuilder("{");
            for (int index = 0; index < keys.size(); index++) {
                if (index > 0) result.append(',');
                String key = keys.get(index);
                result.append(com.fasterxml.jackson.databind.node.TextNode.valueOf(key)).append(':').append(canonicalJson(node.get(key)));
            }
            return result.append('}').toString();
        }
        if (node.isArray()) {
            StringBuilder result = new StringBuilder("[");
            for (int index = 0; index < node.size(); index++) {
                if (index > 0) result.append(',');
                result.append(canonicalJson(node.get(index)));
            }
            return result.append(']').toString();
        }
        return node.toString();
    }

    private static void validateDrawingSourceRange(JsonNode value, java.util.Map<String, int[]> sheetDimensions, String label) {
        if (value == null || !value.isObject()) throw ServiceException.validation(label + " source range is invalid");
        int[] dimensions = sheetDimensions.get(value.path("sheetId").asText());
        for (String coordinate : java.util.List.of("startRow", "endRow", "startColumn", "endColumn")) {
            if (!value.path(coordinate).canConvertToInt() || value.path(coordinate).intValue() < 0) {
                throw ServiceException.validation(label + " source range is invalid");
            }
        }
        int startRow = value.path("startRow").intValue();
        int endRow = value.path("endRow").intValue();
        int startColumn = value.path("startColumn").intValue();
        int endColumn = value.path("endColumn").intValue();
        if (dimensions == null || startRow > endRow || startColumn > endColumn
                || endRow >= dimensions[0] || endColumn >= dimensions[1]) {
            throw ServiceException.validation(label + " source range is outside its worksheet bounds");
        }
        long rows = (long) endRow - startRow + 1;
        long columns = (long) endColumn - startColumn + 1;
        if (rows > GeneratedWorkbookContract.MAX_DRAWING_SOURCE_CELLS || columns > GeneratedWorkbookContract.MAX_DRAWING_SOURCE_CELLS
                || rows * columns > GeneratedWorkbookContract.MAX_DRAWING_SOURCE_CELLS) {
            throw ServiceException.validation(label + " source range exceeds the rendering limit");
        }
    }

    /** One-way migration used only when reading persisted v2 checkpoints. */
    public static ObjectNode migrateStored(JsonNode value, String expectedUnitId) {
        if (value == null || !value.isObject()) throw ServiceException.validation("Stored workbook snapshot must be an object");
        ObjectNode snapshot = ((ObjectNode) value).deepCopy();
        if (snapshot.path("version").asInt(-1) != GeneratedWorkbookContract.SNAPSHOT_VERSION && containsLegacyImageData(snapshot)) {
            throw ServiceException.validation("ASSET_MIGRATION_REQUIRED: legacy image data must be assetized before server persistence");
        }
        if (snapshot.path("version").asInt(-1) == GeneratedWorkbookContract.SNAPSHOT_VERSION) {
            return requireCanonical(snapshot, expectedUnitId);
        }
        if (snapshot.path("version").asInt(-1) == 7 && snapshot.path("sheets").isArray()) {
            snapshot.put("version", GeneratedWorkbookContract.SNAPSHOT_VERSION);
            for (JsonNode raw : (ArrayNode) snapshot.path("sheets")) {
                if (!raw.isObject()) throw ServiceException.validation("Stored workbook snapshot sheet is invalid");
                migrateLegacyReview((ObjectNode) raw);
            }
            return requireCanonical(snapshot, expectedUnitId);
        }
        if (snapshot.path("version").asInt(-1) == 6 && snapshot.path("sheets").isArray()) {
            if (containsLegacyImageData(snapshot)) throw ServiceException.validation("ASSET_MIGRATION_REQUIRED: legacy image data must be assetized before server persistence");
            snapshot.put("version", GeneratedWorkbookContract.SNAPSHOT_VERSION);
            for (JsonNode raw : (ArrayNode) snapshot.path("sheets")) {
                if (!raw.isObject()) throw ServiceException.validation("Stored workbook snapshot sheet is invalid");
                migrateLegacyReview((ObjectNode) raw);
            }
            return requireCanonical(snapshot, expectedUnitId);
        }
        if (snapshot.path("version").asInt(-1) == 5 && snapshot.path("sheets").isArray()) {
            snapshot.put("version", 6);
            return migrateStored(snapshot, expectedUnitId);
        }
        if (snapshot.path("version").asInt(-1) == 4 && snapshot.path("sheets").isArray()) {
            snapshot.put("version", GeneratedWorkbookContract.SNAPSHOT_VERSION);
            ObjectNode dataModel = snapshot.putObject("dataModel");
            dataModel.set("sources", snapshot.path("dataSources").isArray() ? snapshot.path("dataSources").deepCopy() : snapshot.arrayNode());
            dataModel.set("tables", snapshot.path("tables").isArray() ? snapshot.path("tables").deepCopy() : snapshot.arrayNode());
            dataModel.set("relationships", snapshot.arrayNode());
            dataModel.set("views", snapshot.arrayNode());
            snapshot.remove(java.util.List.of("dataSources", "tables"));
            for (JsonNode raw : (ArrayNode) snapshot.path("sheets")) if (raw.isObject()) {
                ObjectNode sheet = (ObjectNode) raw;
                if (!sheet.has("kind")) sheet.put("kind", "worksheet");
                migrateLegacyReview(sheet);
            }
            return requireCanonical(snapshot, expectedUnitId);
        }
        if (snapshot.path("version").asInt(-1) == 3 && snapshot.path("sheets").isArray()) {
            snapshot.put("version", 4);
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
            return migrateStored(snapshot, expectedUnitId);
        }
        if (snapshot.path("version").asInt(-1) != 2 || !snapshot.path("sheets").isArray()) {
            throw ServiceException.validation("Stored workbook snapshot version is invalid");
        }
        snapshot.put("version", 4);
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
        return migrateStored(snapshot, expectedUnitId);
    }

    private static void validateReviewSnapshot(JsonNode value, String sheetId) {
        if (value == null || !value.isObject()) throw ServiceException.validation("Workbook snapshot review is required");
        ObjectNode review = (ObjectNode) value;
        JsonNode notesByCell = review.get("notesByCell");
        JsonNode notesById = review.get("notesById");
        JsonNode threadIdsByCell = review.get("threadIdsByCell");
        JsonNode threadsById = review.get("threadsById");
        if (notesByCell == null || !notesByCell.isObject() || notesById == null || !notesById.isObject()
                || threadIdsByCell == null || !threadIdsByCell.isObject() || threadsById == null || !threadsById.isObject()) {
            throw ServiceException.validation("Workbook snapshot review indexes are required");
        }
        notesById.fields().forEachRemaining(entry -> {
            if (entry.getKey().isBlank() || !entry.getValue().isObject() || !entry.getValue().path("id").isTextual()
                    || entry.getValue().path("id").asText().isBlank() || !entry.getKey().equals(entry.getValue().path("id").asText())) {
                throw ServiceException.validation("Workbook snapshot review note identity is invalid: " + entry.getKey());
            }
        });
        java.util.Set<String> indexedNotes = new java.util.HashSet<>();
        notesByCell.fields().forEachRemaining(entry -> {
            reviewCoordinateKey(entry.getKey());
            if (!entry.getValue().isTextual() || entry.getValue().asText().isBlank() || !notesById.has(entry.getValue().asText()) || !indexedNotes.add(entry.getValue().asText())) {
                throw ServiceException.validation("Workbook snapshot review note index is invalid: " + entry.getKey());
            }
        });
        if (indexedNotes.size() != notesById.size()) throw ServiceException.validation("Workbook snapshot review contains an unindexed note");
        java.util.Set<String> indexedThreads = new java.util.HashSet<>();
        threadIdsByCell.fields().forEachRemaining(entry -> {
            int[] coordinate = reviewCoordinateKey(entry.getKey());
            if (!entry.getValue().isArray()) throw ServiceException.validation("Workbook snapshot review thread index is invalid: " + entry.getKey());
            java.util.Set<String> local = new java.util.HashSet<>();
            for (JsonNode id : entry.getValue()) {
                if (!id.isTextual() || id.asText().isBlank() || !local.add(id.asText()) || !indexedThreads.add(id.asText())) {
                    throw ServiceException.validation("Workbook snapshot review thread index is invalid: " + entry.getKey());
                }
                JsonNode thread = threadsById.get(id.asText());
                if (thread == null || !thread.isObject() || !id.asText().equals(thread.path("id").asText())
                        || !sheetId.equals(thread.path("sheetId").asText()) || thread.path("row").asInt(-1) != coordinate[0]
                        || thread.path("column").asInt(-1) != coordinate[1]) {
                    throw ServiceException.validation("Workbook snapshot review thread index is incompatible: " + id.asText());
                }
            }
        });
        threadsById.fields().forEachRemaining(entry -> {
            JsonNode thread = entry.getValue();
            if (entry.getKey().isBlank() || !thread.isObject() || !entry.getKey().equals(thread.path("id").asText()) || !sheetId.equals(thread.path("sheetId").asText())
                    || !thread.path("row").canConvertToInt() || thread.path("row").asInt(-1) < 0 || thread.path("row").asInt(-1) > 1_048_575
                    || !thread.path("column").canConvertToInt() || thread.path("column").asInt(-1) < 0 || thread.path("column").asInt(-1) > 16_383 || !indexedThreads.contains(entry.getKey())) {
                throw ServiceException.validation("Workbook snapshot review thread identity is invalid: " + entry.getKey());
            }
        });
    }

    private static int[] reviewCoordinateKey(String key) {
        if (key == null || !key.matches("(0|[1-9][0-9]*):(0|[1-9][0-9]*)")) throw ServiceException.validation("Workbook snapshot review cell key is invalid: " + key);
        String[] parts = key.split(":", -1);
        try {
            int row = Integer.parseInt(parts[0]);
            int column = Integer.parseInt(parts[1]);
            if (row < 0 || row > 1_048_575 || column < 0 || column > 16_383) throw ServiceException.validation("Workbook snapshot review cell key is out of bounds: " + key);
            return new int[]{row, column};
        } catch (NumberFormatException exception) {
            throw ServiceException.validation("Workbook snapshot review cell key is invalid: " + key);
        }
    }

    private static void migrateLegacyReview(ObjectNode sheet) {
        ObjectNode review = sheet.objectNode();
        ObjectNode notesByCell = review.putObject("notesByCell");
        ObjectNode notesById = review.putObject("notesById");
        ObjectNode threadIdsByCell = review.putObject("threadIdsByCell");
        ObjectNode threadsById = review.putObject("threadsById");
        JsonNode legacyNotes = sheet.get("notes");
        if (legacyNotes != null && !legacyNotes.isNull() && !legacyNotes.isArray()) throw ServiceException.validation("Legacy worksheet notes must be an array");
        if (legacyNotes != null && legacyNotes.isArray()) for (JsonNode raw : legacyNotes) {
            if (!raw.isObject() || !raw.path("note").isObject()) throw ServiceException.validation("Legacy worksheet note is invalid");
            int row = raw.path("row").asInt(-1);
            int column = raw.path("column").asInt(-1);
            putMigratedNote(notesByCell, notesById, row, column, raw.get("note"));
        }
        JsonNode legacyThreads = sheet.get("commentThreads");
        if (legacyThreads != null && !legacyThreads.isNull() && !legacyThreads.isArray()) throw ServiceException.validation("Legacy worksheet comments must be an array");
        if (legacyThreads != null && legacyThreads.isArray()) for (JsonNode raw : legacyThreads) putMigratedThread(sheet.path("id").asText(), threadsById, threadIdsByCell, raw, -1, -1);
        JsonNode cells = sheet.get("cells");
        if (cells != null && cells.isObject()) cells.fields().forEachRemaining(rowEntry -> {
            int row = parseLegacyCoordinate(rowEntry.getKey(), "row");
            if (!rowEntry.getValue().isObject()) throw ServiceException.validation("Legacy cell row is invalid");
            rowEntry.getValue().fields().forEachRemaining(columnEntry -> {
                int column = parseLegacyCoordinate(columnEntry.getKey(), "column");
                if (!columnEntry.getValue().isObject()) throw ServiceException.validation("Legacy cell is invalid");
                ObjectNode cell = (ObjectNode) columnEntry.getValue();
                if (cell.has("note")) putMigratedNote(notesByCell, notesById, row, column, cell.get("note"));
                if (cell.has("comment")) putMigratedThread(sheet.path("id").asText(), threadsById, threadIdsByCell, cell.get("comment"), row, column);
                cell.remove(java.util.List.of("note", "comment"));
            });
        });
        sheet.set("review", review);
        sheet.remove(java.util.List.of("notes", "commentThreads"));
    }

    private static int parseLegacyCoordinate(String value, String label) {
        try {
            int parsed = Integer.parseInt(value);
            if (parsed < 0) throw ServiceException.validation("Legacy " + label + " is invalid");
            return parsed;
        } catch (NumberFormatException exception) {
            throw ServiceException.validation("Legacy " + label + " is invalid");
        }
    }

    private static void putMigratedNote(ObjectNode notesByCell, ObjectNode notesById, int row, int column, JsonNode value) {
        if (row < 0 || row > 1_048_575 || column < 0 || column > 16_383 || value == null || !value.isObject() || !value.path("id").isTextual() || value.path("id").asText().isBlank()) {
            throw ServiceException.validation("Legacy worksheet note is invalid");
        }
        String key = row + ":" + column;
        String id = value.path("id").asText();
        if (notesByCell.has(key) && (!id.equals(notesByCell.path(key).asText()) || !canonicalJson(notesById.get(id)).equals(canonicalJson(value)))) {
            throw ServiceException.validation("REVIEW_MIGRATION_CONFLICT: note at " + key);
        }
        notesByCell.fields().forEachRemaining(entry -> {
            if (id.equals(entry.getValue().asText()) && !key.equals(entry.getKey())) throw ServiceException.validation("REVIEW_MIGRATION_CONFLICT: note identity " + id);
        });
        if (notesById.has(id) && !canonicalJson(notesById.get(id)).equals(canonicalJson(value))) throw ServiceException.validation("REVIEW_MIGRATION_CONFLICT: note identity " + id);
        notesByCell.put(key, id);
        notesById.set(id, value.deepCopy());
    }

    private static void putMigratedThread(String sheetId, ObjectNode threadsById, ObjectNode threadIdsByCell, JsonNode value, int row, int column) {
        if (value == null || !value.isObject() || !value.path("id").isTextual() || value.path("id").asText().isBlank()) throw ServiceException.validation("Legacy worksheet comment is invalid");
        ObjectNode thread = (ObjectNode) value.deepCopy();
        if (thread.has("sheetId") && !sheetId.equals(thread.path("sheetId").asText())) throw ServiceException.validation("REVIEW_MIGRATION_CONFLICT: comment targets another sheet");
        thread.put("sheetId", sheetId);
        if (row >= 0) thread.put("row", row);
        if (column >= 0) thread.put("column", column);
        if (!thread.path("row").canConvertToInt() || !thread.path("column").canConvertToInt() || thread.path("row").asInt(-1) < 0 || thread.path("column").asInt(-1) < 0) throw ServiceException.validation("Legacy worksheet comment location is invalid");
        if (!thread.has("replies")) thread.putArray("replies");
        String id = thread.path("id").asText();
        if (threadsById.has(id)) {
            if (!canonicalJson(threadsById.get(id)).equals(canonicalJson(thread))) throw ServiceException.validation("REVIEW_MIGRATION_CONFLICT: comment identity " + id);
            return;
        }
        threadsById.set(id, thread);
        String key = thread.path("row").asInt() + ":" + thread.path("column").asInt();
        ArrayNode ids = threadIdsByCell.withArray(key);
        for (JsonNode indexedId : ids) if (id.equals(indexedId.asText())) return;
        ids.add(id);
    }

    private static String canonicalJson(JsonNode value) {
        return value == null ? "null" : value.toString();
    }

    private static boolean containsLegacyImageData(JsonNode value) {
        if (value == null) return false;
        if (value.isArray()) {
            for (JsonNode entry : value) if (containsLegacyImageData(entry)) return true;
            return false;
        }
        if (!value.isObject()) return false;
        if ("image".equals(value.path("kind").asText()) && value.path("src").isTextual()) return true;
        var fields = value.fields();
        while (fields.hasNext()) if (containsLegacyImageData(fields.next().getValue())) return true;
        return false;
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
        for (String coordinate : java.util.List.of("startRow", "endRow", "startColumn", "endColumn")) {
            if (!value.has(coordinate) || !value.get(coordinate).isIntegralNumber()) {
                throw ServiceException.validation("AutoFilter range coordinates are required");
            }
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
