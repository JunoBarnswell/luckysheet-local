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
    private static final int MAX_ROW_INDEX = 1_048_575;
    private static final int MAX_COLUMN_INDEX = 16_383;
    private static final java.util.Set<String> PANE_NONE_FIELDS = java.util.Set.of("kind");
    private static final java.util.Set<String> PANE_FIELDS = java.util.Set.of("kind", "state", "xSplit", "ySplit", "startRow", "startColumn", "activePane");
    private static final java.util.Set<String> HYPERLINK_ENTRY_FIELDS = java.util.Set.of("row", "column", "hyperlink");
    private static final java.util.Set<String> HYPERLINK_FIELDS = java.util.Set.of("id", "target", "tooltip");
    private static final java.util.Set<String> HYPERLINK_URL_FIELDS = java.util.Set.of("kind", "url");
    private static final java.util.Set<String> HYPERLINK_EMAIL_FIELDS = java.util.Set.of("kind", "address", "subject");
    private static final java.util.Set<String> HYPERLINK_SHEET_ADDRESS_FIELDS = java.util.Set.of("kind", "sheetId", "address");
    private static final java.util.Set<String> HYPERLINK_SHEET_COORDINATE_FIELDS = java.util.Set.of("kind", "sheetId", "row", "column");
    private static final java.util.Set<String> HYPERLINK_NAME_FIELDS = java.util.Set.of("kind", "name");
    private static final java.util.Set<String> HYPERLINK_URL_SCHEMES = java.util.Set.of("http", "https", "ftp");
    private static final java.util.regex.Pattern HYPERLINK_EMAIL_ADDRESS = java.util.regex.Pattern.compile("^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$");
    private static final java.util.regex.Pattern HYPERLINK_SHEET_ADDRESS = java.util.regex.Pattern.compile("^([A-Za-z]+)([1-9][0-9]*)$");
    private static final java.util.regex.Pattern HYPERLINK_DEFINED_NAME = java.util.regex.Pattern.compile("^[A-Za-z_\\\\][A-Za-z0-9_.]*$");

    private WorkbookSnapshotValidator() {
    }

    public static void requireCanonicalPane(JsonNode pane) {
        if (pane == null || !pane.isObject()) throw ServiceException.validation("Workbook snapshot pane is invalid");
        String kind = pane.path("kind").asText();
        if ("none".equals(kind)) {
            requirePaneFields(pane, PANE_NONE_FIELDS);
            if (pane.has("state") || pane.has("xSplit") || pane.has("ySplit")
                    || pane.has("startRow") || pane.has("startColumn") || pane.has("activePane")) {
                throw ServiceException.validation("Workbook snapshot none pane contains split state");
            }
            return;
        }
        if (!("frozen".equals(kind) || "split".equals(kind))) {
            throw ServiceException.validation("Workbook snapshot pane kind is invalid");
        }
        requirePaneFields(pane, PANE_FIELDS);
        String state = pane.path("state").asText();
        if (("frozen".equals(kind) && !("frozen".equals(state) || "frozenSplit".equals(state)))
                || ("split".equals(kind) && !"split".equals(state))) {
            throw ServiceException.validation("Workbook snapshot pane state is invalid");
        }
        requirePaneCoordinate(pane, "startRow", MAX_ROW_INDEX);
        requirePaneCoordinate(pane, "startColumn", MAX_COLUMN_INDEX);
        if ("frozen".equals(kind)) {
            requireFrozenSplit(pane, "xSplit", MAX_COLUMN_INDEX + 1);
            requireFrozenSplit(pane, "ySplit", MAX_ROW_INDEX + 1);
        } else {
            requireSplitPosition(pane, "xSplit");
            requireSplitPosition(pane, "ySplit");
        }
        JsonNode activePane = pane.get("activePane");
        if (activePane != null && (!activePane.isTextual()
                || !java.util.Set.of("topLeft", "topRight", "bottomLeft", "bottomRight").contains(activePane.asText()))) {
            throw ServiceException.validation("Workbook snapshot pane activePane is invalid");
        }
    }

    private static void requirePaneFields(JsonNode pane, java.util.Set<String> allowedFields) {
        pane.fieldNames().forEachRemaining(field -> {
            if (!allowedFields.contains(field)) throw ServiceException.validation("Workbook snapshot pane field is not canonical: " + field);
        });
    }

    private static void requirePaneCoordinate(JsonNode pane, String field, int maximum) {
        JsonNode coordinate = pane.get(field);
        if (coordinate == null || !coordinate.isIntegralNumber() || !coordinate.canConvertToInt()
                || coordinate.intValue() < 0 || coordinate.intValue() > maximum) {
            throw ServiceException.validation("Workbook snapshot pane " + field + " is invalid");
        }
    }

    private static void requireFrozenSplit(JsonNode pane, String field, int maximum) {
        JsonNode split = pane.get(field);
        if (split == null || !split.isIntegralNumber() || !split.canConvertToInt()
                || split.intValue() < 0 || split.intValue() > maximum) {
            throw ServiceException.validation("Workbook snapshot frozen pane " + field + " is invalid");
        }
    }

    private static void requireSplitPosition(JsonNode pane, String field) {
        JsonNode split = pane.get(field);
        if (split == null || !split.isNumber() || !Double.isFinite(split.asDouble()) || split.asDouble() < 0) {
            throw ServiceException.validation("Workbook snapshot split pane " + field + " is invalid");
        }
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
                    || !sheet.path("drawings").isArray() || !sheet.path("drawingPayloads").isObject()
                    || !sheet.path("hyperlinks").isArray()) {
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
                    || !sheet.path("pane").isObject()) {
                throw ServiceException.validation("Workbook snapshot sheet pixel geometry is invalid");
            }
            requireCanonicalPane(sheet.path("pane"));
            sheet.path("drawingPayloads").fields().forEachRemaining(entry -> {
                JsonNode payload = entry.getValue();
                if ("camera".equals(payload.path("kind").asText())) {
                    validateDrawingSourceRange(payload.get("sourceRange"), sheetDimensions, "Camera");
                }
                if ("image".equals(payload.path("kind").asText())) validateAssetRef(payload.get("asset"), "Drawing image");
            });
            sheet.path("cells").fields().forEachRemaining(row -> row.getValue().fields().forEachRemaining(cell -> {
                if (cell.getValue().has("hyperlink") || cell.getValue().has("hyperlinkDetail")) {
                    throw ServiceException.validation("Workbook snapshot cell contains legacy hyperlink metadata");
                }
                JsonNode presentation = cell.getValue().get("presentation");
                if (presentation != null && "image".equals(presentation.path("kind").asText())) validateAssetRef(presentation.get("asset"), "Cell image");
            }));
            validateCellBounds(sheet);
            validateCanonicalHyperlinks(sheet, snapshot, sheetId, sheetDimensions);
            JsonNode autoFilter = sheet.get("autoFilter");
            if (autoFilter != null && !autoFilter.isNull()) validateAutoFilter(autoFilter, sheetId, null);
            JsonNode tables = sheet.get("sheetTables");
            if (tables != null && !tables.isNull()) {
                if (!tables.isArray()) throw ServiceException.validation("Workbook snapshot sheetTables is invalid");
                for (JsonNode table : tables) {
                    if (!table.isObject()) throw ServiceException.validation("Workbook snapshot table is invalid");
                    RangeRef tableRange = rangeOf(table.get("range"), sheetId);
                    JsonNode columns = table.get("columns");
                    int tableWidth = tableRange.endColumn() - tableRange.startColumn() + 1;
                    if (columns == null || !columns.isArray() || columns.size() != tableWidth) {
                        throw ServiceException.validation("Workbook snapshot Sheet Table columns must match its range width");
                    }
                    JsonNode tableFilter = table.get("autoFilter");
                    if (tableFilter == null || tableFilter.isNull()) continue;
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
                JsonNode pivotSource = "chart".equals(kind) ? payload.get("source") : payload;
                if (pivotSource == null || !pivotSource.isObject()) {
                    throw ServiceException.validation("Drawing " + drawingId + " is missing its canonical Pivot source");
                }
                if ("chart".equals(kind) && !"pivot".equals(pivotSource.path("kind").asText())) continue;
                String pivotId = "chart".equals(kind) ? pivotSource.path("pivotId").asText() : payload.path("pivotId").asText();
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
        return snapshot;
    }

    private static void validateCellBounds(JsonNode sheet) {
        int rowCount = sheet.path("rowCount").intValue();
        int columnCount = sheet.path("columnCount").intValue();
        sheet.path("cells").fields().forEachRemaining(row -> {
            int rowIndex;
            try {
                rowIndex = Integer.parseInt(row.getKey());
            } catch (NumberFormatException error) {
                throw ServiceException.validation("Workbook snapshot cell row is invalid");
            }
            if (rowIndex < 0 || rowIndex >= rowCount || !row.getValue().isObject()) {
                throw ServiceException.validation("Workbook snapshot cell is outside worksheet bounds");
            }
            row.getValue().fields().forEachRemaining(cell -> {
                int columnIndex;
                try {
                    columnIndex = Integer.parseInt(cell.getKey());
                } catch (NumberFormatException error) {
                    throw ServiceException.validation("Workbook snapshot cell column is invalid");
                }
                if (columnIndex < 0 || columnIndex >= columnCount || !cell.getValue().isObject()) {
                    throw ServiceException.validation("Workbook snapshot cell is outside worksheet bounds");
                }
            });
        });
    }

    private static void validateCanonicalHyperlinks(JsonNode sheet, JsonNode snapshot, String sourceSheetId,
                                                     java.util.Map<String, int[]> sheetDimensions) {
        java.util.Set<String> positions = new java.util.HashSet<>();
        for (JsonNode entry : sheet.path("hyperlinks")) {
            if (!entry.isObject() || !hasOnlyFields(entry, HYPERLINK_ENTRY_FIELDS)
                    || !entry.path("row").isIntegralNumber() || !entry.path("row").canConvertToInt()
                    || !entry.path("column").isIntegralNumber() || !entry.path("column").canConvertToInt()) {
                throw ServiceException.validation("Workbook snapshot hyperlink coordinate is invalid");
            }
            int row = entry.path("row").intValue();
            int column = entry.path("column").intValue();
            int[] dimensions = sheetDimensions.get(sourceSheetId);
            String position = row + ":" + column;
            if (dimensions == null || row < 0 || column < 0 || row >= dimensions[0] || column >= dimensions[1]) {
                throw ServiceException.validation("Workbook snapshot hyperlink is outside worksheet bounds");
            }
            if (!positions.add(position)) throw ServiceException.validation("Workbook snapshot hyperlink coordinate is duplicated");
            JsonNode hyperlink = entry.get("hyperlink");
            if (hyperlink == null || !hyperlink.isObject() || !hyperlink.path("id").isTextual() || hyperlink.path("id").asText().isBlank()
                    || !hasOnlyFields(hyperlink, HYPERLINK_FIELDS)
                    || !hyperlink.path("target").isObject()
                    || (hyperlink.has("tooltip") && !hyperlink.get("tooltip").isTextual())) {
                throw ServiceException.validation("Workbook snapshot hyperlink is invalid");
            }
            validateCanonicalHyperlinkTarget(hyperlink.get("target"), snapshot, sourceSheetId, sheetDimensions);
        }
    }

    private static void validateCanonicalHyperlinkTarget(JsonNode target, JsonNode snapshot, String sourceSheetId,
                                                          java.util.Map<String, int[]> sheetDimensions) {
        if (!target.path("kind").isTextual()) throw ServiceException.validation("Workbook snapshot hyperlink target kind is invalid");
        String kind = target.path("kind").asText();
        switch (kind) {
            case "url" -> {
                if (!hasOnlyFields(target, HYPERLINK_URL_FIELDS)) throw ServiceException.validation("Workbook snapshot hyperlink URL target has unsupported fields");
                String value = target.path("url").asText().trim();
                try {
                    java.net.URI uri = new java.net.URI(value);
                    String scheme = uri.getScheme();
                    if (value.isBlank() || scheme == null
                            || !HYPERLINK_URL_SCHEMES.contains(scheme.toLowerCase(java.util.Locale.ROOT))
                            || uri.getHost() == null || uri.getHost().isBlank()) {
                        throw ServiceException.validation("Workbook snapshot hyperlink URL is invalid");
                    }
                } catch (java.net.URISyntaxException exception) {
                    throw ServiceException.validation("Workbook snapshot hyperlink URL is invalid");
                }
            }
            case "email" -> {
                if (!hasOnlyFields(target, HYPERLINK_EMAIL_FIELDS)) throw ServiceException.validation("Workbook snapshot email hyperlink target has unsupported fields");
                String address = target.path("address").asText().trim();
                if (!HYPERLINK_EMAIL_ADDRESS.matcher(address).matches()
                        || (target.has("subject") && !target.get("subject").isTextual())) {
                    throw ServiceException.validation("Workbook snapshot email hyperlink is invalid");
                }
            }
            case "sheet" -> {
                if (!target.path("sheetId").isTextual()) throw ServiceException.validation("Workbook snapshot hyperlink target sheet is invalid");
                String targetSheetId = target.path("sheetId").asText();
                int[] dimensions = sheetDimensions.get(targetSheetId);
                if (targetSheetId.isBlank() || dimensions == null) {
                    throw ServiceException.validation("Workbook snapshot hyperlink target sheet does not exist");
                }
                boolean hasAddress = target.has("address");
                boolean hasRow = target.has("row");
                boolean hasColumn = target.has("column");
                if (!hasOnlyFields(target, hasAddress
                        ? HYPERLINK_SHEET_ADDRESS_FIELDS
                        : HYPERLINK_SHEET_COORDINATE_FIELDS)) {
                    throw ServiceException.validation("Workbook snapshot worksheet hyperlink target has unsupported fields");
                }
                if (hasAddress && (hasRow || hasColumn) || !hasAddress && !(hasRow && hasColumn)) {
                    throw ServiceException.validation("Workbook snapshot worksheet hyperlink address is invalid");
                }
                int row;
                int column;
                if (hasAddress) {
                    java.util.regex.Matcher matcher = HYPERLINK_SHEET_ADDRESS.matcher(target.path("address").asText().trim());
                    if (!matcher.matches()) throw ServiceException.validation("Workbook snapshot worksheet hyperlink address is invalid");
                    column = columnIndex(matcher.group(1));
                    try {
                        row = Math.subtractExact(Integer.parseInt(matcher.group(2)), 1);
                    } catch (NumberFormatException | ArithmeticException exception) {
                        throw ServiceException.validation("Workbook snapshot worksheet hyperlink row is invalid");
                    }
                } else {
                    if (!target.path("row").isIntegralNumber() || !target.path("row").canConvertToInt()
                            || !target.path("column").isIntegralNumber() || !target.path("column").canConvertToInt()) {
                        throw ServiceException.validation("Workbook snapshot worksheet hyperlink coordinates are invalid");
                    }
                    row = target.path("row").intValue();
                    column = target.path("column").intValue();
                }
                if (row < 0 || column < 0 || row >= dimensions[0] || column >= dimensions[1]) {
                    throw ServiceException.validation("Workbook snapshot worksheet hyperlink is outside worksheet bounds");
                }
            }
            case "name" -> {
                if (!hasOnlyFields(target, HYPERLINK_NAME_FIELDS)) throw ServiceException.validation("Workbook snapshot defined-name hyperlink target has unsupported fields");
                if (!target.path("name").isTextual()) throw ServiceException.validation("Workbook snapshot defined-name hyperlink is invalid");
                String name = target.path("name").asText();
                if (!HYPERLINK_DEFINED_NAME.matcher(name).matches() || !hasHyperlinkName(snapshot, sourceSheetId, name)) {
                    throw ServiceException.validation("Workbook snapshot defined-name hyperlink is invalid");
                }
            }
            default -> throw ServiceException.validation("Workbook snapshot hyperlink target kind is invalid");
        }
    }

    private static int columnIndex(String label) {
        long value = 0;
        for (int index = 0; index < label.length(); index++) {
            int digit = Character.toUpperCase(label.charAt(index)) - 'A' + 1;
            if (digit < 1 || digit > 26) throw ServiceException.validation("Workbook snapshot worksheet hyperlink column is invalid");
            try {
                value = Math.addExact(Math.multiplyExact(value, 26), digit);
            } catch (ArithmeticException exception) {
                throw ServiceException.validation("Workbook snapshot worksheet hyperlink column is invalid");
            }
        }
        if (value < 1 || value > Integer.MAX_VALUE) throw ServiceException.validation("Workbook snapshot worksheet hyperlink column is invalid");
        return (int) value - 1;
    }

    private static boolean hasHyperlinkName(JsonNode snapshot, String sourceSheetId, String name) {
        JsonNode models = snapshot.get("definedNameModels");
        if (models != null) {
            if (!models.isArray()) return false;
            for (JsonNode entry : models) {
                if (name.equalsIgnoreCase(entry.path("name").asText())
                        && ("workbook".equals(entry.path("scope").asText())
                        || "sheet".equals(entry.path("scope").asText()) && sourceSheetId.equals(entry.path("sheetId").asText()))) return true;
            }
            return false;
        }
        JsonNode projection = snapshot.get("definedNames");
        if (projection != null && projection.isObject()) {
            java.util.Iterator<String> names = projection.fieldNames();
            while (names.hasNext()) if (name.equalsIgnoreCase(names.next())) return true;
        }
        return false;
    }

    private static boolean hasOnlyFields(JsonNode object, java.util.Set<String> allowed) {
        java.util.Iterator<String> fields = object.fieldNames();
        while (fields.hasNext()) if (!allowed.contains(fields.next())) return false;
        return true;
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

    public static String canonicalJson(JsonNode node) {
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
                if (!Integer.toString(key).equals(entry.getKey())
                        || !column.isObject() || column.path("column").asInt(Integer.MIN_VALUE) != key
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

}
