package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

/** Strict JSON primitives shared by server-side snapshot reducers. */
final class SnapshotMutationSupport {
    static final int MAX_ROW = 1_048_575;
    static final int MAX_COLUMN = 16_383;
    static final int MAX_CHANGED_CELLS = 100_000;

    private SnapshotMutationSupport() {
    }

    static ObjectNode root(JsonNode snapshot) {
        if (snapshot == null || !snapshot.isObject()) throw ServiceException.validation("Workbook snapshot must be an object");
        ObjectNode root = (ObjectNode) snapshot;
        sheets(root);
        return root;
    }

    static ArrayNode sheets(ObjectNode root) {
        JsonNode sheets = root.get("sheets");
        if (sheets == null || !sheets.isArray()) throw ServiceException.validation("Snapshot sheets array is required");
        return (ArrayNode) sheets;
    }

    static ObjectNode params(OperationMutation mutation) {
        if (mutation.params() == null || !mutation.params().isObject()) throw ServiceException.validation("Mutation params must be an object");
        ObjectNode params = (ObjectNode) mutation.params();
        JsonNode declaredSheet = params.get("sheetId");
        if (declaredSheet != null && (!declaredSheet.isTextual() || !mutation.sheetId().equals(declaredSheet.asText()))) {
            throw ServiceException.validation("Mutation sheetId does not match params");
        }
        return params;
    }

    static ObjectNode sheet(ObjectNode root, String id) {
        if (id == null || id.isBlank()) throw ServiceException.validation("Mutation sheetId is required");
        for (JsonNode candidate : sheets(root)) {
            if (candidate.isObject() && id.equals(candidate.path("id").asText())) return (ObjectNode) candidate;
        }
        throw ServiceException.notFound("Sheet not found: " + id);
    }

    static ObjectNode object(ObjectNode parent, String property) {
        JsonNode value = parent.get(property);
        if (value == null || value.isNull()) {
            ObjectNode created = JsonNodeFactory.instance.objectNode();
            parent.set(property, created);
            return created;
        }
        if (!value.isObject()) throw ServiceException.validation(property + " must be an object");
        return (ObjectNode) value;
    }

    static ObjectNode requiredObject(ObjectNode parent, String property) {
        JsonNode value = parent.get(property);
        if (value == null || !value.isObject()) throw ServiceException.validation(property + " must be an object");
        return (ObjectNode) value;
    }

    static ArrayNode array(ObjectNode parent, String property) {
        JsonNode value = parent.get(property);
        if (value == null || value.isNull()) {
            ArrayNode created = JsonNodeFactory.instance.arrayNode();
            parent.set(property, created);
            return created;
        }
        if (!value.isArray()) throw ServiceException.validation(property + " must be an array");
        return (ArrayNode) value;
    }

    static ArrayNode dataModelArray(ObjectNode root, String property) {
        return array(requiredObject(root, "dataModel"), property);
    }

    static ArrayNode requiredArray(ObjectNode parent, String property) {
        JsonNode value = parent.get(property);
        if (value == null || !value.isArray()) throw ServiceException.validation(property + " must be an array");
        return (ArrayNode) value;
    }

    static ObjectNode cells(ObjectNode sheet) {
        return object(sheet, "cells");
    }

    static ObjectNode cellRow(ObjectNode cells, int row, boolean create) {
        JsonNode value = cells.get(Integer.toString(row));
        if (value == null) {
            if (!create) return null;
            ObjectNode created = cells.objectNode();
            cells.set(Integer.toString(row), created);
            return created;
        }
        if (!value.isObject()) throw ServiceException.validation("Cell row must be an object");
        return (ObjectNode) value;
    }

    static ObjectNode cell(ObjectNode sheet, CellCoordinate coordinate, boolean create) {
        ObjectNode row = cellRow(cells(sheet), coordinate.row(), create);
        if (row == null) return null;
        JsonNode current = row.get(Integer.toString(coordinate.column()));
        if (current == null) {
            if (!create) return null;
            ObjectNode cell = row.objectNode();
            cell.putNull("value");
            row.set(Integer.toString(coordinate.column()), cell);
            return cell;
        }
        if (!current.isObject()) throw ServiceException.validation("Cell value must be an object");
        return (ObjectNode) current;
    }

    static void putCell(ObjectNode sheet, CellCoordinate coordinate, JsonNode value) {
        cellRow(cells(sheet), coordinate.row(), true).set(Integer.toString(coordinate.column()), value.deepCopy());
    }

    static void removeCell(ObjectNode sheet, CellCoordinate coordinate) {
        ObjectNode cells = cells(sheet);
        ObjectNode row = cellRow(cells, coordinate.row(), false);
        if (row == null) return;
        row.remove(Integer.toString(coordinate.column()));
        if (row.isEmpty()) cells.remove(Integer.toString(coordinate.row()));
    }

    static void clearCells(ObjectNode sheet, RangeRef range) {
        ObjectNode cells = cells(sheet);
        for (int row = range.startRow(); row <= range.endRow(); row++) {
            ObjectNode current = cellRow(cells, row, false);
            if (current == null) continue;
            List<String> remove = new ArrayList<>();
            current.fieldNames().forEachRemaining(key -> {
                try {
                    int column = Integer.parseInt(key);
                    if (column >= range.startColumn() && column <= range.endColumn()) remove.add(key);
                } catch (NumberFormatException ignored) {
                    throw ServiceException.validation("Cell column key is invalid");
                }
            });
            remove.forEach(current::remove);
            if (current.isEmpty()) cells.remove(Integer.toString(row));
        }
    }

    static void removeHyperlinks(ObjectNode sheet, RangeRef range) {
        ArrayNode hyperlinks = array(sheet, "hyperlinks");
        for (int index = hyperlinks.size() - 1; index >= 0; index--) {
            JsonNode hyperlink = hyperlinks.get(index);
            int row = hyperlink.path("row").asInt(-1);
            int column = hyperlink.path("column").asInt(-1);
            if (row >= range.startRow() && row <= range.endRow() && column >= range.startColumn() && column <= range.endColumn()) {
                hyperlinks.remove(index);
            }
        }
    }

    static CellCoordinate coordinate(ObjectNode root, String sheetId, ObjectNode params) {
        sheet(root, sheetId);
        int row = index(root, sheetId, params, "row");
        int column = index(root, sheetId, params, "column");
        return new CellCoordinate(row, column);
    }

    static int index(ObjectNode root, String sheetId, ObjectNode params, String property) {
        JsonNode value = params.get(property);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt()) throw ServiceException.validation(property + " must be an integer");
        int index = value.intValue();
        int maximum = "row".equals(property) ? MAX_ROW : MAX_COLUMN;
        if (index < 0 || index > maximum) throw ServiceException.validation(property + " is out of bounds");
        sheet(root, sheetId);
        return index;
    }

    static RangeRef cellRange(ObjectNode root, String sheetId, ObjectNode params) {
        CellCoordinate coordinate = coordinate(root, sheetId, params);
        return new RangeRef(sheetId, coordinate.row(), coordinate.row(), coordinate.column(), coordinate.column());
    }

    static RangeRef rowRange(ObjectNode root, String sheetId, ObjectNode params) {
        int row = index(root, sheetId, params, "row");
        ObjectNode sheet = sheet(root, sheetId);
        return new RangeRef(sheetId, row, row, 0, canonicalDimension(sheet, "columnCount") - 1);
    }

    static RangeRef wholeSheetRange(ObjectNode root, String sheetId) {
        ObjectNode sheet = sheet(root, sheetId);
        int rowCount = canonicalDimension(sheet, "rowCount");
        int columnCount = canonicalDimension(sheet, "columnCount");
        return new RangeRef(sheetId, 0, rowCount - 1, 0, columnCount - 1);
    }

    static RangeRef columnRange(ObjectNode root, String sheetId, ObjectNode params) {
        int column = index(root, sheetId, params, "column");
        ObjectNode sheet = sheet(root, sheetId);
        return new RangeRef(sheetId, 0, canonicalDimension(sheet, "rowCount") - 1, column, column);
    }

    private static int canonicalDimension(ObjectNode sheet, String field) {
        JsonNode value = sheet.get(field);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 1) {
            throw ServiceException.validation("Canonical worksheet " + field + " is required");
        }
        return value.intValue();
    }

    static RangeRef range(ObjectNode root, JsonNode value) {
        if (value == null || !value.isObject()) throw ServiceException.validation("Range must be an object");
        ObjectNode range = (ObjectNode) value;
        String sheetId = text(range, "sheetId");
        sheet(root, sheetId);
        int startRow = rangeIndex(range, "startRow", MAX_ROW);
        int endRow = rangeIndex(range, "endRow", MAX_ROW);
        int startColumn = rangeIndex(range, "startColumn", MAX_COLUMN);
        int endColumn = rangeIndex(range, "endColumn", MAX_COLUMN);
        if (endRow < startRow || endColumn < startColumn) throw ServiceException.validation("Range bounds are invalid");
        return new RangeRef(sheetId, startRow, endRow, startColumn, endColumn);
    }

    static List<RangeRef> styleRanges(ObjectNode root, String sheetId, ObjectNode params) {
        List<RangeRef> ranges = new ArrayList<>();
        if (params.get("range") != null) ranges.add(range(root, params.get("range")));
        JsonNode many = params.get("ranges");
        if (many != null) {
            if (!many.isArray()) throw ServiceException.validation("ranges must be an array");
            for (JsonNode range : many) ranges.add(range(root, range));
        }
        if (ranges.isEmpty()) throw ServiceException.validation("A style mutation requires a range");
        for (RangeRef range : ranges) requireSheet(range, sheetId);
        long count = ranges.stream().mapToLong(SnapshotMutationSupport::cellCount).sum();
        if (count > MAX_CHANGED_CELLS) throw ServiceException.validation("Style range is too large");
        return List.copyOf(ranges);
    }

    static Matrix matrix(ObjectNode root, String sheetId, ObjectNode params) {
        int startRow = index(root, sheetId, params, "startRow");
        int startColumn = index(root, sheetId, params, "startColumn");
        JsonNode values = params.get("values");
        if (values == null || !values.isArray()) throw ServiceException.validation("Range values must be an array");
        ArrayNode matrix = (ArrayNode) values;
        if (matrix.size() == 0) return new Matrix(startRow, startColumn, List.of());
        if (matrix.size() > MAX_CHANGED_CELLS) throw ServiceException.validation("Range values are too large");
        List<ArrayNode> rows = new ArrayList<>();
        int largestWidth = 0;
        for (JsonNode row : matrix) {
            if (!row.isArray()) throw ServiceException.validation("Range values must contain rows");
            if (row.size() > MAX_CHANGED_CELLS) throw ServiceException.validation("Range row is too large");
            rows.add((ArrayNode) row);
            largestWidth = Math.max(largestWidth, row.size());
        }
        long count = (long) rows.size() * largestWidth;
        if (count > MAX_CHANGED_CELLS) throw ServiceException.validation("Range values are too large");
        if (startRow + rows.size() - 1 > MAX_ROW || startColumn + largestWidth - 1 > MAX_COLUMN) throw ServiceException.validation("Range values exceed worksheet bounds");
        return new Matrix(startRow, startColumn, List.copyOf(rows));
    }

    static RangeRef matrixRange(ObjectNode root, String sheetId, ObjectNode params) {
        Matrix matrix = matrix(root, sheetId, params);
        if (matrix.values().isEmpty() || matrix.width() == 0) return null;
        return new RangeRef(sheetId, matrix.startRow(), matrix.startRow() + matrix.values().size() - 1, matrix.startColumn(), matrix.startColumn() + matrix.width() - 1);
    }

    static void requireSheet(RangeRef range, String sheetId) {
        if (!sheetId.equals(range.sheetId())) throw ServiceException.validation("Range sheetId does not match mutation sheetId");
    }

    static boolean contains(RangeRef range, CellCoordinate coordinate) {
        return coordinate.row() >= range.startRow() && coordinate.row() <= range.endRow()
                && coordinate.column() >= range.startColumn() && coordinate.column() <= range.endColumn();
    }

    static long cellCount(RangeRef range) {
        return (long) (range.endRow() - range.startRow() + 1) * (range.endColumn() - range.startColumn() + 1);
    }

    static void removeNotes(ObjectNode sheet, RangeRef range) {
        ArrayNode notes = array(sheet, "notes");
        for (int index = notes.size() - 1; index >= 0; index--) {
            JsonNode entry = notes.get(index);
            if (entry.isObject() && contains(range, new CellCoordinate(entry.path("row").asInt(-1), entry.path("column").asInt(-1)))) notes.remove(index);
        }
    }

    static boolean removeNote(ArrayNode notes, CellCoordinate coordinate) {
        for (int index = notes.size() - 1; index >= 0; index--) {
            JsonNode entry = notes.get(index);
            if (entry.path("row").asInt(-1) == coordinate.row() && entry.path("column").asInt(-1) == coordinate.column()) {
                notes.remove(index);
                return true;
            }
        }
        return false;
    }

    static ObjectNode findNote(ArrayNode notes, CellCoordinate coordinate) {
        for (JsonNode entry : notes) {
            if (entry.isObject() && entry.path("row").asInt(-1) == coordinate.row() && entry.path("column").asInt(-1) == coordinate.column()) return (ObjectNode) entry;
        }
        return null;
    }

    static void removeThreads(ObjectNode sheet, RangeRef range) {
        ArrayNode threads = array(sheet, "commentThreads");
        for (int index = threads.size() - 1; index >= 0; index--) {
            JsonNode thread = threads.get(index);
            if (thread.isObject() && contains(range, new CellCoordinate(thread.path("row").asInt(-1), thread.path("column").asInt(-1)))) threads.remove(index);
        }
    }

    static void restoreNotes(ObjectNode root, ObjectNode sheet, String sheetId, RangeRef range, JsonNode value) {
        if (value == null || value.isNull()) return;
        if (!value.isArray()) throw ServiceException.validation("Range restore notes must be an array");
        if (value.size() > MAX_CHANGED_CELLS) throw ServiceException.validation("Range restore notes are too large");
        ArrayNode notes = array(sheet, "notes");
        for (JsonNode entry : value) {
            if (!entry.isObject()) throw ServiceException.validation("Range restore note must be an object");
            CellCoordinate coordinate = coordinate(root, sheetId, (ObjectNode) entry);
            if (!contains(range, coordinate)) throw ServiceException.validation("Range restore note is outside its range");
            JsonNode note = entry.get("note");
            if (note == null || !note.isObject()) throw ServiceException.validation("Range restore note payload is invalid");
            removeNote(notes, coordinate);
            ObjectNode next = notes.objectNode();
            next.put("row", coordinate.row());
            next.put("column", coordinate.column());
            next.set("note", note.deepCopy());
            notes.add(next);
        }
    }

    static void restoreThreads(ObjectNode root, ObjectNode sheet, String sheetId, RangeRef range, JsonNode value) {
        if (value == null || value.isNull()) return;
        if (!value.isArray()) throw ServiceException.validation("Range restore comments must be an array");
        if (value.size() > MAX_CHANGED_CELLS) throw ServiceException.validation("Range restore comments are too large");
        ArrayNode threads = array(sheet, "commentThreads");
        for (JsonNode thread : value) {
            if (!thread.isObject()) throw ServiceException.validation("Range restore comment must be an object");
            if (!sheetId.equals(thread.path("sheetId").asText())) throw ServiceException.validation("Range restore comment targets another sheet");
            CellCoordinate coordinate = coordinate(root, sheetId, (ObjectNode) thread);
            if (!contains(range, coordinate)) throw ServiceException.validation("Range restore comment is outside its range");
            String id = text((ObjectNode) thread, "id");
            ArrayNode next = JsonNodeFactory.instance.arrayNode();
            for (JsonNode existing : threads) if (!id.equals(existing.path("id").asText())) next.add(existing);
            threads.removeAll();
            threads.addAll(next);
            threads.add(thread.deepCopy());
        }
    }

    static ObjectNode findThread(ArrayNode threads, String id) {
        for (JsonNode thread : threads) if (thread.isObject() && id.equals(thread.path("id").asText())) return (ObjectNode) thread;
        return null;
    }

    static ObjectNode requireThread(ObjectNode sheet, ObjectNode params) {
        String id = text(params, "threadId");
        ObjectNode thread = findThread(array(sheet, "commentThreads"), id);
        if (thread == null) throw ServiceException.notFound("Comment thread not found");
        return thread;
    }

    static RangeRef threadRange(ObjectNode root, String sheetId, ObjectNode params) {
        ObjectNode sheet = sheet(root, sheetId);
        ObjectNode thread = requireThread(sheet, params);
        ObjectNode location = JsonNodeFactory.instance.objectNode();
        location.put("row", thread.path("row").asInt(-1));
        location.put("column", thread.path("column").asInt(-1));
        return cellRange(root, sheetId, location);
    }

    static String text(ObjectNode object, String property) {
        JsonNode value = object.get(property);
        if (value == null || !value.isTextual() || value.asText().isBlank()) throw ServiceException.validation(property + " is required");
        return value.asText();
    }

    static String optionalText(ObjectNode object, String property) {
        JsonNode value = object.get(property);
        if (value == null || value.isNull()) return null;
        if (!value.isTextual()) throw ServiceException.validation(property + " must be a string");
        return value.asText();
    }

    static void requireEntitySheet(ObjectNode entity, String expectedSheetId) {
        String actual = text(entity, "sheetId");
        if (!expectedSheetId.equals(actual)) throw ServiceException.validation("Entity sheetId does not match mutation sheetId");
    }

    static ObjectNode findById(ArrayNode entries, String id) {
        for (JsonNode entry : entries) {
            if (entry.isObject() && id.equals(entry.path("id").asText())) return (ObjectNode) entry;
        }
        return null;
    }

    static int indexById(ArrayNode entries, String id) {
        for (int index = 0; index < entries.size(); index++) {
            if (entries.get(index).isObject() && id.equals(entries.get(index).path("id").asText())) return index;
        }
        return -1;
    }

    static void upsertById(ArrayNode entries, ObjectNode entity) {
        String id = text(entity, "id");
        int index = indexById(entries, id);
        if (index >= 0) entries.set(index, entity.deepCopy());
        else entries.add(entity.deepCopy());
    }

    static ObjectNode requireById(ArrayNode entries, String id, String entityName) {
        ObjectNode entity = findById(entries, id);
        if (entity == null) throw ServiceException.notFound(entityName + " not found: " + id);
        return entity;
    }

    static boolean removeById(ArrayNode entries, String id) {
        int index = indexById(entries, id);
        if (index < 0) return false;
        entries.remove(index);
        return true;
    }

    static List<RangeRef> ranges(ObjectNode root, JsonNode values, String expectedSheetId) {
        if (values == null || !values.isArray()) throw ServiceException.validation("ranges must be an array");
        List<RangeRef> ranges = new ArrayList<>();
        for (JsonNode value : values) {
            RangeRef range = range(root, value);
            requireSheet(range, expectedSheetId);
            ranges.add(range);
        }
        if (ranges.isEmpty()) throw ServiceException.validation("At least one range is required");
        return List.copyOf(ranges);
    }

    static void validateKnownKeys(ObjectNode object, Set<String> allowed, String label) {
        object.fieldNames().forEachRemaining(key -> {
            if (!allowed.contains(key)) throw ServiceException.validation(label + " contains unknown field: " + key);
        });
    }

    static boolean sameAnchor(JsonNode rangeNode, RangeRef range) {
        return rangeNode != null && rangeNode.isObject()
                && range.sheetId().equals(rangeNode.path("sheetId").asText())
                && range.startRow() == rangeNode.path("startRow").asInt(-1)
                && range.startColumn() == rangeNode.path("startColumn").asInt(-1);
    }

    static ObjectNode rangeNode(RangeRef range, ArrayNode owner) {
        ObjectNode node = owner.objectNode();
        node.put("sheetId", range.sheetId());
        node.put("startRow", range.startRow());
        node.put("endRow", range.endRow());
        node.put("startColumn", range.startColumn());
        node.put("endColumn", range.endColumn());
        return node;
    }

    private static int rangeIndex(ObjectNode range, String property, int maximum) {
        JsonNode value = range.get(property);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 0 || value.intValue() > maximum) {
            throw ServiceException.validation(property + " is out of bounds");
        }
        return value.intValue();
    }

    private static int positiveDimension(JsonNode value, String property, int fallback) {
        if (value == null || value.isMissingNode() || value.isNull()) return fallback;
        if (!value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 1 || value.intValue() > fallback) {
            throw ServiceException.validation(property + " is invalid");
        }
        return value.intValue();
    }

    record CellCoordinate(int row, int column) {
    }

    record Matrix(int startRow, int startColumn, List<ArrayNode> values) {
        int width() {
            return values.stream().mapToInt(ArrayNode::size).max().orElse(0);
        }
    }
}
