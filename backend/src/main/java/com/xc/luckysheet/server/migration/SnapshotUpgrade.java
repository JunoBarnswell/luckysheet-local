package com.xc.luckysheet.server.migration;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.xc.luckysheet.server.contract.GeneratedWorkbookContract;
import com.xc.luckysheet.server.service.ServiceException;
import static com.xc.luckysheet.server.contract.WorkbookSnapshotValidator.requireCanonical;
import static com.xc.luckysheet.server.contract.WorkbookSnapshotValidator.canonicalJson;
/** Only invoked at an explicit persisted-data migration boundary. */
public final class SnapshotUpgrade {
    private SnapshotUpgrade() { }
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

    private static ObjectNode copyObjectOrEmpty(JsonNode value) {
        return value != null && value.isObject() ? ((ObjectNode) value).deepCopy() : com.fasterxml.jackson.databind.node.JsonNodeFactory.instance.objectNode();
    }

    private static void migrateFontSizes(JsonNode value) {
        if (value == null) return;
        if (value.isArray()) value.forEach(SnapshotUpgrade::migrateFontSizes);
        if (!value.isObject()) return;
        ObjectNode object = (ObjectNode) value;
        if (object.path("fontSize").isNumber() && !object.has("fontSizePx")) object.set("fontSizePx", object.get("fontSize"));
        object.remove("fontSize");
        java.util.List<JsonNode> children = new java.util.ArrayList<>();
        object.elements().forEachRemaining(children::add);
        children.forEach(SnapshotUpgrade::migrateFontSizes);
    }
}
