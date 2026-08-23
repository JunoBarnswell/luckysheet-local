package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

/**
 * Canonical structural reducer for a workbook JSON snapshot.
 *
 * It owns the same participants as the browser structural transform: cells,
 * ranges, table sources, objects, notes/reviews, draw anchors, pivot and
 * sparkline definitions, protections, names, and formula references. It
 * never receives a replacement snapshot from a client.
 */
final class StructuralSnapshotReducer {
    private StructuralSnapshotReducer() {
    }

    static void applyAxis(
            ObjectNode root,
            String sheetId,
            FormulaReferenceTransformer.Axis axis,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction
    ) {
        PivotMutationDescriptor.assertCanonicalSnapshot(root);
        ObjectNode target = SnapshotMutationSupport.sheet(root, sheetId);
        int limit = dimension(target, axis);
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW + 1 : SnapshotMutationSupport.MAX_COLUMN + 1;
        validateAxisBounds(limit, maximum, at, count, direction);
        if (direction == FormulaReferenceTransformer.Direction.DELETE) validateDeletePreservation(root, target, axis, at, count);

        remapCells(target, axis, at, count, direction);
        setDimension(target, axis, direction == FormulaReferenceTransformer.Direction.INSERT ? limit + count : Math.max(1, limit - count));
        shiftAllMetadata(root, target, sheetId, axis, at, count, direction);
        rewriteAxisFormulas(root, target, axis, at, count, direction);
    }

    static void shiftCells(ObjectNode root, String sheetId, RangeRef source, ShiftDirection direction) {
        PivotMutationDescriptor.assertCanonicalSnapshot(root);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        SnapshotMutationSupport.requireSheet(source, sheetId);
        RangeRef selection = normalize(source);
        int rowDelta = direction == ShiftDirection.DOWN ? 1 : direction == ShiftDirection.UP ? -1 : 0;
        int columnDelta = direction == ShiftDirection.RIGHT ? 1 : direction == ShiftDirection.LEFT ? -1 : 0;
        validateShiftPreservation(root, sheet, selection, rowDelta, columnDelta);

        List<CellEntry> sourceCells = cellsInRange(sheet, selection);
        SnapshotMutationSupport.clearCells(sheet, selection);
        for (CellEntry entry : sourceCells) {
            int nextRow = entry.row() + rowDelta;
            int nextColumn = entry.column() + columnDelta;
            if (!contains(selection, nextRow, nextColumn)) continue;
            ObjectNode cell = entry.cell().deepCopy();
            if (cell.path("formula").isTextual()) {
                cell.put("formula", FormulaReferenceTransformer.offset(cell.path("formula").asText(), rowDelta, columnDelta));
                cell.remove("formulaValue");
            }
            SnapshotMutationSupport.putCell(sheet, new SnapshotMutationSupport.CellCoordinate(nextRow, nextColumn), cell);
        }
        shiftBoundedMetadata(root, sheet, sheetId, selection, rowDelta, columnDelta);
        rewriteMovedReferences(root, sheet, sheetId, selection, rowDelta, columnDelta);
    }

    static void restoreShiftedCells(ObjectNode root, String sheetId, RangeRef range, ShiftDirection originalDirection, JsonNode cells) {
        if (originalDirection != null) shiftCells(root, sheetId, range, originalDirection.reverse());
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        RangeRef normalized = normalize(range);
        SnapshotMutationSupport.clearCells(sheet, normalized);
        if (cells == null || !cells.isArray()) throw ServiceException.validation("Structural restore cells must be an array");
        if (cells.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Structural restore is too large");
        for (JsonNode raw : cells) {
            if (!raw.isObject()) throw ServiceException.validation("Structural restore cell must be an object");
            ObjectNode entry = (ObjectNode) raw;
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, entry);
            if (!contains(normalized, coordinate.row(), coordinate.column())) throw ServiceException.validation("Structural restore cell is outside range");
            JsonNode cell = entry.get("cell");
            if (cell == null || !cell.isObject()) throw ServiceException.validation("Structural restore cell payload must be an object");
            SnapshotMutationSupport.putCell(sheet, coordinate, cell);
        }
    }

    static void permuteRows(ObjectNode root, String sheetId, RangeRef range, JsonNode sourceRows) {
        PivotMutationDescriptor.assertCanonicalSnapshot(root);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        SnapshotMutationSupport.requireSheet(range, sheetId);
        RangeRef selected = normalize(range);
        if (sourceRows == null || !sourceRows.isArray()) throw ServiceException.validation("Row permutation sourceRows must be an array");
        int expected = selected.endRow() - selected.startRow() + 1;
        if (sourceRows.size() != expected) throw ServiceException.validation("Row permutation length does not match range");
        validatePermutationPreservation(sheet, selected);
        int[] mapping = validatePermutation(selected, (ArrayNode) sourceRows);
        remapPermutedCells(sheet, selected, mapping);
        remapPermutationMetadata(sheet, selected, mapping);
    }

    private static void validateAxisBounds(int limit, int maximum, int at, int count, FormulaReferenceTransformer.Direction direction) {
        if (at < 0 || count < 1) throw ServiceException.validation("Structural bounds are invalid");
        if (direction == FormulaReferenceTransformer.Direction.INSERT) {
            if (at > limit || (long) limit + count > maximum) throw ServiceException.validation("Structural insert exceeds worksheet bounds");
            return;
        }
        if (at >= limit || at + count > limit) throw ServiceException.validation("Structural delete is outside worksheet bounds");
    }

    private static int dimension(ObjectNode sheet, FormulaReferenceTransformer.Axis axis) {
        String property = axis == FormulaReferenceTransformer.Axis.ROW ? "rowCount" : "columnCount";
        JsonNode value = sheet.get(property);
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW + 1 : SnapshotMutationSupport.MAX_COLUMN + 1;
        if (value == null || !value.isIntegralNumber() || value.intValue() < 1 || value.intValue() > maximum) throw ServiceException.validation(property + " is invalid");
        return value.intValue();
    }

    private static void setDimension(ObjectNode sheet, FormulaReferenceTransformer.Axis axis, int value) {
        int maximum = axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW + 1 : SnapshotMutationSupport.MAX_COLUMN + 1;
        if (value < 1 || value > maximum) throw ServiceException.validation("Structural result exceeds worksheet bounds");
        sheet.put(axis == FormulaReferenceTransformer.Axis.ROW ? "rowCount" : "columnCount", value);
    }

    private static void remapCells(ObjectNode sheet, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ObjectNode oldCells = SnapshotMutationSupport.cells(sheet);
        ObjectNode next = JsonNodeFactory.instance.objectNode();
        for (java.util.Iterator<java.util.Map.Entry<String, JsonNode>> rows = oldCells.fields(); rows.hasNext();) {
            java.util.Map.Entry<String, JsonNode> rowEntry = rows.next();
            int row = integerKey(rowEntry.getKey(), SnapshotMutationSupport.MAX_ROW, "Cell row");
            if (!rowEntry.getValue().isObject()) throw ServiceException.validation("Cell row must be an object");
            for (java.util.Iterator<java.util.Map.Entry<String, JsonNode>> columns = ((ObjectNode) rowEntry.getValue()).fields(); columns.hasNext();) {
                java.util.Map.Entry<String, JsonNode> columnEntry = columns.next();
                int column = integerKey(columnEntry.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                if (!columnEntry.getValue().isObject()) throw ServiceException.validation("Cell payload must be an object");
                int nextRow = axis == FormulaReferenceTransformer.Axis.ROW ? shiftIndex(row, at, count, direction) : row;
                int nextColumn = axis == FormulaReferenceTransformer.Axis.COLUMN ? shiftIndex(column, at, count, direction) : column;
                if (nextRow < 0 || nextColumn < 0) continue;
                ObjectNode rowTarget = next.with(Integer.toString(nextRow));
                rowTarget.set(Integer.toString(nextColumn), columnEntry.getValue().deepCopy());
            }
        }
        sheet.set("cells", next);
    }

    private static int shiftIndex(int value, int at, int count, FormulaReferenceTransformer.Direction direction) {
        if (direction == FormulaReferenceTransformer.Direction.INSERT) return value >= at ? value + count : value;
        if (value < at) return value;
        if (value < at + count) return -1;
        return value - count;
    }

    private static int integerKey(String value, int maximum, String label) {
        try {
            int parsed = Integer.parseInt(value);
            if (parsed < 0 || parsed > maximum) throw ServiceException.validation(label + " is out of bounds");
            return parsed;
        } catch (NumberFormatException error) {
            throw ServiceException.validation(label + " is invalid");
        }
    }

    private static RangeRef normalize(RangeRef range) {
        return new RangeRef(range.sheetId(), Math.min(range.startRow(), range.endRow()), Math.max(range.startRow(), range.endRow()), Math.min(range.startColumn(), range.endColumn()), Math.max(range.startColumn(), range.endColumn()));
    }

    private static boolean contains(RangeRef range, int row, int column) {
        return row >= range.startRow() && row <= range.endRow() && column >= range.startColumn() && column <= range.endColumn();
    }

    private static void validateDeletePreservation(ObjectNode root, ObjectNode target, FormulaReferenceTransformer.Axis axis, int at, int count) {
        int end = at + count - 1;
        for (JsonNode sparkline : SnapshotMutationSupport.array(target, "sparklines")) {
            if (insideDeleted(sparkline.path("anchor").path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a sparkline anchor");
            }
        }
        for (JsonNode pivot : SnapshotMutationSupport.array(target, "pivots")) {
            ObjectNode pivotTarget = PivotMutationDescriptor.requiredTarget(pivot);
            JsonNode anchor = pivotTarget.path("anchor");
            if (anchor.isObject() && insideDeleted(anchor.path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a pivot target anchor");
            }
        }
        for (JsonNode spill : SnapshotMutationSupport.array(target, "spillRanges")) {
            if (insideDeleted(spill.path("anchor").path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a spill anchor");
            }
        }
        for (JsonNode drawing : SnapshotMutationSupport.array(target, "drawings")) {
            JsonNode anchor = drawing.path("anchor");
            if (anchor.path("kind").asText().equals("absolute")) continue;
            String startKey = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            String endKey = axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn";
            if (insideDeleted(anchor.path(startKey).asInt(-1), at, end) || insideDeleted(anchor.path(endKey).asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a drawing anchor");
            }
        }
        for (JsonNode note : SnapshotMutationSupport.array(target, "notes")) {
            if (insideDeleted(note.path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a note");
            }
        }
        for (JsonNode comment : SnapshotMutationSupport.array(target, "commentThreads")) {
            if (insideDeleted(comment.path(axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column").asInt(-1), at, end)) {
                throw ServiceException.validation("Structural delete would lose a comment thread");
            }
        }
        for (JsonNode table : SnapshotMutationSupport.array(target, "sheetTables")) {
            if (intersectsAxis(table.get("range"), axis, at, count)) throw ServiceException.validation("Structural delete intersects a sheet table");
        }
        String targetSheetId = target.path("id").asText();
        for (JsonNode table : SnapshotMutationSupport.array(root, "tables")) {
            JsonNode range = table.get("sourceRange");
            if (range != null && targetSheetId.equals(range.path("sheetId").asText()) && intersectsAxis(range, axis, at, count)) {
                throw ServiceException.validation("Structural delete intersects a workbook table");
            }
        }
    }

    private static boolean insideDeleted(int position, int at, int end) {
        return position >= at && position <= end;
    }

    private static boolean intersectsAxis(JsonNode range, FormulaReferenceTransformer.Axis axis, int at, int count) {
        if (range == null || !range.isObject()) throw ServiceException.validation("Structural participant range is invalid");
        String start = axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn";
        String end = axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn";
        return range.path(start).asInt(Integer.MAX_VALUE) <= at + count - 1 && range.path(end).asInt(-1) >= at;
    }

    private static void shiftAllMetadata(
            ObjectNode root,
            ObjectNode target,
            String targetSheetId,
            FormulaReferenceTransformer.Axis axis,
            int at,
            int count,
            FormulaReferenceTransformer.Direction direction
    ) {
        shiftMerges(root, target, targetSheetId, axis, at, count, direction);
        shiftFreeze(target, axis, at, count, direction);
        shiftHiddenAndSizes(target, axis, at, count, direction);
        shiftTargetNotes(target, axis, at, count, direction);
        shiftTargetComments(target, axis, at, count, direction);
        shiftTargetDrawings(target, axis, at, count, direction);
        shiftTargetSpills(root, target, targetSheetId, axis, at, count, direction);
        shiftTargetSheetTables(root, target, targetSheetId, axis, at, count, direction);
        shiftTargetProtectionAndOutline(root, target, targetSheetId, axis, at, count, direction);

        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            if (!raw.isObject()) throw ServiceException.validation("Workbook contains an invalid sheet");
            ObjectNode owner = (ObjectNode) raw;
            shiftRules(root, owner, targetSheetId, "conditionalFormats", axis, at, count, direction);
            shiftRules(root, owner, targetSheetId, "dataValidations", axis, at, count, direction);
            shiftFilter(root, owner, targetSheetId, axis, at, count, direction);
            shiftPivots(root, owner, targetSheetId, axis, at, count, direction);
            shiftSparklines(root, owner, targetSheetId, axis, at, count, direction);
            shiftChartPayloads(root, owner, targetSheetId, axis, at, count, direction);
        }
        shiftWorkbookTables(root, targetSheetId, axis, at, count, direction);
    }

    private static boolean shiftRange(ObjectNode root, JsonNode raw, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Structural participant range is invalid");
        ObjectNode range = (ObjectNode) raw;
        SnapshotMutationSupport.range(root, range);
        if (!targetSheetId.equals(range.path("sheetId").asText())) return true;
        String startKey = axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn";
        String endKey = axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn";
        int start = range.path(startKey).asInt();
        int end = range.path(endKey).asInt();
        if (direction == FormulaReferenceTransformer.Direction.INSERT) {
            if (start >= at) {
                start += count;
                end += count;
            } else if (end >= at) end += count;
            range.put(startKey, start);
            range.put(endKey, end);
            return end >= start;
        }
        int deletedEnd = at + count - 1;
        if (end < at) return true;
        if (start > deletedEnd) {
            range.put(startKey, start - count);
            range.put(endKey, end - count);
            return true;
        }
        end = Math.max(at - 1, end - count);
        start = Math.min(Math.max(at - 1, start), end);
        range.put(startKey, start);
        range.put(endKey, end);
        return end >= start;
    }

    private static void shiftMerges(ObjectNode root, ObjectNode sheet, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ArrayNode merges = SnapshotMutationSupport.array(sheet, "merges");
        for (int index = merges.size() - 1; index >= 0; index--) {
            ObjectNode merge = requireObject(merges.get(index), "Merge");
            boolean keep = shiftRange(root, merge.get("range"), targetSheetId, axis, at, count, direction);
            if (!keep) {
                merges.remove(index);
                continue;
            }
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(merge, "anchor");
            String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            int current = anchor.path(key).asInt(-1);
            if (direction == FormulaReferenceTransformer.Direction.INSERT && current >= at) anchor.put(key, current + count);
            if (direction == FormulaReferenceTransformer.Direction.DELETE) {
                if (current >= at && current < at + count) anchor.put(key, merge.path("range").path(axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn").asInt());
                else if (current >= at + count) anchor.put(key, current - count);
            }
        }
    }

    private static void shiftRules(ObjectNode root, ObjectNode owner, String targetSheetId, String property, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(owner, property)) {
            ObjectNode rule = requireObject(raw, "Range rule");
            ArrayNode ranges = SnapshotMutationSupport.array(rule, "ranges");
            ArrayNode next = JsonNodeFactory.instance.arrayNode();
            for (JsonNode range : ranges) if (shiftRange(root, range, targetSheetId, axis, at, count, direction)) next.add(range);
            rule.set("ranges", next);
        }
    }

    private static void shiftFilter(ObjectNode root, ObjectNode owner, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        JsonNode filterRaw = owner.get("filter");
        if (filterRaw == null || filterRaw.isNull()) return;
        ObjectNode filter = requireObject(filterRaw, "Filter");
        if (!shiftRange(root, filter.get("range"), targetSheetId, axis, at, count, direction)) {
            owner.remove("filter");
            return;
        }
        if (axis != FormulaReferenceTransformer.Axis.COLUMN || !targetSheetId.equals(owner.path("id").asText())) return;
        ObjectNode criteria = SnapshotMutationSupport.requiredObject(filter, "criteria");
        ObjectNode next = JsonNodeFactory.instance.objectNode();
        criteria.fields().forEachRemaining(entry -> {
            int column = integerKey(entry.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Filter criteria column");
            int shifted = shiftIndex(column, at, count, direction);
            if (shifted < 0) return;
            ObjectNode condition = requireObject(entry.getValue(), "Filter condition").deepCopy();
            condition.put("column", shifted);
            next.set(Integer.toString(shifted), condition);
        });
        filter.set("criteria", next);
    }

    private static void shiftFreeze(ObjectNode sheet, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ObjectNode freeze = SnapshotMutationSupport.requiredObject(sheet, "pane");
        if ("none".equals(freeze.path("kind").asText())) return;
        String split = axis == FormulaReferenceTransformer.Axis.ROW ? "ySplit" : "xSplit";
        String start = axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn";
        List<String> keys = "frozen".equals(freeze.path("kind").asText()) ? List.of(split, start) : List.of(start);
        for (String key : keys) {
            int value = freeze.path(key).asInt(0);
            if (direction == FormulaReferenceTransformer.Direction.INSERT && value >= at) freeze.put(key, value + count);
            if (direction == FormulaReferenceTransformer.Direction.DELETE && value > at) freeze.put(key, Math.max(0, value - count));
        }
    }

    private static void shiftHiddenAndSizes(ObjectNode sheet, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        String hidden = axis == FormulaReferenceTransformer.Axis.ROW ? "hiddenRows" : "hiddenColumns";
        String sizes = axis == FormulaReferenceTransformer.Axis.ROW ? "rowHeightsPx" : "columnWidthsPx";
        ArrayNode remapped = JsonNodeFactory.instance.arrayNode();
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, hidden)) {
            if (!raw.isIntegralNumber()) throw ServiceException.validation(hidden + " contains invalid index");
            int shifted = shiftIndex(raw.intValue(), at, count, direction);
            if (shifted >= 0 && !containsNumber(remapped, shifted)) remapped.add(shifted);
        }
        sheet.set(hidden, remapped);
        ObjectNode nextSizes = JsonNodeFactory.instance.objectNode();
        SnapshotMutationSupport.object(sheet, sizes).fields().forEachRemaining(entry -> {
            int shifted = shiftIndex(integerKey(entry.getKey(), axis == FormulaReferenceTransformer.Axis.ROW ? SnapshotMutationSupport.MAX_ROW : SnapshotMutationSupport.MAX_COLUMN, sizes), at, count, direction);
            if (shifted >= 0) nextSizes.set(Integer.toString(shifted), entry.getValue());
        });
        sheet.set(sizes, nextSizes);
    }

    private static void shiftSparklines(ObjectNode root, ObjectNode owner, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ArrayNode sparklines = SnapshotMutationSupport.array(owner, "sparklines");
        boolean ownerTarget = targetSheetId.equals(owner.path("id").asText());
        for (int index = sparklines.size() - 1; index >= 0; index--) {
            ObjectNode sparkline = requireObject(sparklines.get(index), "Sparkline");
            shiftRange(root, sparkline.get("sourceRange"), targetSheetId, axis, at, count, direction);
            if (!ownerTarget) continue;
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(sparkline, "anchor");
            String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            int shifted = shiftIndex(anchor.path(key).asInt(-1), at, count, direction);
            if (shifted < 0) sparklines.remove(index);
            else anchor.put(key, shifted);
        }
    }

    private static void shiftPivots(ObjectNode root, ObjectNode owner, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        boolean ownerTarget = targetSheetId.equals(owner.path("id").asText());
        for (JsonNode raw : SnapshotMutationSupport.array(owner, "pivots")) {
            ObjectNode pivot = requireObject(raw, "Pivot");
            SnapshotMutationSupport.validateKnownKeys(pivot, Set.of("schema", "id", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "nativeMetadata"), "Pivot");
            PivotMutationDescriptor.forEachWorksheetSourceRange(pivot, range -> shiftRange(root, range, targetSheetId, axis, at, count, direction));
            ObjectNode target = PivotMutationDescriptor.requiredTarget(pivot);
            JsonNode anchorRaw = target.get("anchor");
            if (ownerTarget && anchorRaw != null && anchorRaw.isObject()) {
                ObjectNode anchor = (ObjectNode) anchorRaw;
                String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
                int shifted = shiftIndex(anchor.path(key).asInt(-1), at, count, direction);
                if (shifted >= 0) anchor.put(key, shifted);
            }
        }
    }

    private static void shiftChartPayloads(ObjectNode root, ObjectNode owner, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ObjectNode payloads = SnapshotMutationSupport.object(owner, "drawingPayloads");
        payloads.fields().forEachRemaining(entry -> {
            JsonNode raw = entry.getValue();
            if (!raw.isObject() || !"chart".equals(raw.path("kind").asText())) return;
            ObjectNode chart = (ObjectNode) raw;
            for (JsonNode range : chart.path("sourceRanges")) shiftRange(root, range, targetSheetId, axis, at, count, direction);
            if (chart.has("categoryRange")) shiftRange(root, chart.get("categoryRange"), targetSheetId, axis, at, count, direction);
            for (JsonNode series : chart.path("series")) {
                if (series.isObject() && series.has("range")) shiftRange(root, series.get("range"), targetSheetId, axis, at, count, direction);
                if (series.isObject() && series.has("xRange")) shiftRange(root, series.get("xRange"), targetSheetId, axis, at, count, direction);
                if (series.isObject() && series.has("yRange")) shiftRange(root, series.get("yRange"), targetSheetId, axis, at, count, direction);
            }
        });
    }

    private static void shiftTargetDrawings(ObjectNode target, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "drawings")) {
            ObjectNode drawing = requireObject(raw, "Drawing");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(drawing, "anchor");
            if ("absolute".equals(anchor.path("kind").asText())) continue;
            shiftAnchor(axis, at, count, direction, anchor, axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column");
            String end = axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn";
            if (anchor.has(end)) shiftAnchor(axis, at, count, direction, anchor, end);
        }
    }

    private static void shiftAnchor(FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction, ObjectNode anchor, String property) {
        int shifted = shiftIndex(anchor.path(property).asInt(-1), at, count, direction);
        if (shifted < 0) anchor.put(property, at);
        else anchor.put(property, shifted);
    }

    private static void shiftTargetNotes(ObjectNode target, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "notes")) {
            ObjectNode note = requireObject(raw, "Note");
            String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            int shifted = shiftIndex(note.path(key).asInt(-1), at, count, direction);
            if (shifted < 0) throw ServiceException.validation("Structural delete would lose a note");
            note.put(key, shifted);
        }
    }

    private static void shiftTargetComments(ObjectNode target, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(target, "commentThreads")) {
            ObjectNode comment = requireObject(raw, "Comment thread");
            String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            int shifted = shiftIndex(comment.path(key).asInt(-1), at, count, direction);
            if (shifted < 0) throw ServiceException.validation("Structural delete would lose a comment thread");
            comment.put(key, shifted);
        }
    }

    private static void shiftTargetSpills(ObjectNode root, ObjectNode target, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ArrayNode spills = SnapshotMutationSupport.array(target, "spillRanges");
        for (int index = spills.size() - 1; index >= 0; index--) {
            ObjectNode spill = requireObject(spills.get(index), "Spill range");
            if (!shiftRange(root, spill.get("range"), targetSheetId, axis, at, count, direction)) {
                spills.remove(index);
                continue;
            }
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(spill, "anchor");
            String key = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            int shifted = shiftIndex(anchor.path(key).asInt(-1), at, count, direction);
            if (shifted < 0) spills.remove(index);
            else anchor.put(key, shifted);
        }
    }

    private static void shiftTargetSheetTables(ObjectNode root, ObjectNode target, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ArrayNode tables = SnapshotMutationSupport.array(target, "sheetTables");
        for (int index = tables.size() - 1; index >= 0; index--) {
            if (!shiftRange(root, requireObject(tables.get(index), "Sheet table").get("range"), targetSheetId, axis, at, count, direction)) tables.remove(index);
        }
    }

    private static void shiftTargetProtectionAndOutline(ObjectNode root, ObjectNode target, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        ArrayNode protections = SnapshotMutationSupport.array(target, "protectionRules");
        for (int index = protections.size() - 1; index >= 0; index--) {
            ObjectNode rule = requireObject(protections.get(index), "Protection rule");
            if (rule.has("range") && !shiftRange(root, rule.get("range"), targetSheetId, axis, at, count, direction)) protections.remove(index);
        }
        JsonNode banded = target.get("bandedRule");
        if (banded != null && banded.isObject() && !shiftRange(root, banded.get("range"), targetSheetId, axis, at, count, direction)) target.remove("bandedRule");
        JsonNode outlineRaw = target.get("outline");
        if (outlineRaw == null || !outlineRaw.isObject()) return;
        ArrayNode groups = SnapshotMutationSupport.array((ObjectNode) outlineRaw, "groups");
        for (int index = groups.size() - 1; index >= 0; index--) {
            ObjectNode group = requireObject(groups.get(index), "Outline group");
            String expectedAxis = axis == FormulaReferenceTransformer.Axis.ROW ? "row" : "column";
            if (!expectedAxis.equals(group.path("axis").asText())) continue;
            ObjectNode range = JsonNodeFactory.instance.objectNode();
            range.put("sheetId", targetSheetId);
            if (axis == FormulaReferenceTransformer.Axis.ROW) {
                range.put("startRow", group.path("start").asInt());
                range.put("endRow", group.path("end").asInt());
                range.put("startColumn", 0);
                range.put("endColumn", 0);
            } else {
                range.put("startRow", 0);
                range.put("endRow", 0);
                range.put("startColumn", group.path("start").asInt());
                range.put("endColumn", group.path("end").asInt());
            }
            if (!shiftRange(root, range, targetSheetId, axis, at, count, direction)) {
                groups.remove(index);
                continue;
            }
            group.put("start", range.path(axis == FormulaReferenceTransformer.Axis.ROW ? "startRow" : "startColumn").asInt());
            group.put("end", range.path(axis == FormulaReferenceTransformer.Axis.ROW ? "endRow" : "endColumn").asInt());
        }
    }

    private static void shiftWorkbookTables(ObjectNode root, String targetSheetId, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        for (JsonNode raw : SnapshotMutationSupport.array(root, "tables")) {
            ObjectNode table = requireObject(raw, "Workbook table");
            if (table.has("sourceRange")) shiftRange(root, table.get("sourceRange"), targetSheetId, axis, at, count, direction);
        }
    }

    private static void rewriteAxisFormulas(ObjectNode root, ObjectNode targetSheet, FormulaReferenceTransformer.Axis axis, int at, int count, FormulaReferenceTransformer.Direction direction) {
        FormulaReferenceTransformer.SheetIdentity target = identity(targetSheet);
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(raw, "Sheet");
            FormulaReferenceTransformer.SheetIdentity ownerIdentity = identity(owner);
            forEachFormulaCell(owner, cell -> {
                String original = cell.path("formula").asText();
                String rewritten = FormulaReferenceTransformer.remapAxis(original, ownerIdentity, target, axis, at, count, direction);
                if (!original.equals(rewritten)) {
                    cell.put("formula", rewritten);
                    cell.remove("formulaValue");
                }
            });
        }
        ObjectNode names = SnapshotMutationSupport.object(root, "definedNames");
        names.fields().forEachRemaining(entry -> {
            if (!entry.getValue().isTextual()) return;
            String rewritten = FormulaReferenceTransformer.remapAxis(entry.getValue().asText(), target, target, axis, at, count, direction);
            names.put(entry.getKey(), rewritten);
        });
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode name = requireObject(raw, "Defined name");
            if ("sheet".equals(name.path("scope").asText()) && !target.id().equals(name.path("sheetId").asText())) continue;
            if (!name.path("formula").isTextual()) continue;
            String rewritten = FormulaReferenceTransformer.remapAxis(name.path("formula").asText(), target, target, axis, at, count, direction);
            name.put("formula", rewritten);
        }
    }

    private static FormulaReferenceTransformer.SheetIdentity identity(ObjectNode sheet) {
        return new FormulaReferenceTransformer.SheetIdentity(SnapshotMutationSupport.text(sheet, "id"), SnapshotMutationSupport.text(sheet, "name"));
    }

    private static void forEachFormulaCell(ObjectNode sheet, java.util.function.Consumer<ObjectNode> consumer) {
        ObjectNode cells = SnapshotMutationSupport.cells(sheet);
        cells.fields().forEachRemaining(row -> {
            if (!row.getValue().isObject()) throw ServiceException.validation("Cell row must be an object");
            ((ObjectNode) row.getValue()).fields().forEachRemaining(column -> {
                if (!column.getValue().isObject()) throw ServiceException.validation("Cell payload must be an object");
                ObjectNode cell = (ObjectNode) column.getValue();
                if (cell.path("formula").isTextual()) consumer.accept(cell);
            });
        });
    }

    private static boolean containsNumber(ArrayNode values, int number) {
        for (JsonNode value : values) if (value.isIntegralNumber() && value.intValue() == number) return true;
        return false;
    }

    private static List<CellEntry> cellsInRange(ObjectNode sheet, RangeRef range) {
        List<CellEntry> entries = new ArrayList<>();
        ObjectNode cells = SnapshotMutationSupport.cells(sheet);
        cells.fields().forEachRemaining(row -> {
            int rowIndex = integerKey(row.getKey(), SnapshotMutationSupport.MAX_ROW, "Cell row");
            if (!row.getValue().isObject()) throw ServiceException.validation("Cell row must be an object");
            ((ObjectNode) row.getValue()).fields().forEachRemaining(column -> {
                int columnIndex = integerKey(column.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                if (!column.getValue().isObject()) throw ServiceException.validation("Cell payload must be an object");
                if (contains(range, rowIndex, columnIndex)) entries.add(new CellEntry(rowIndex, columnIndex, ((ObjectNode) column.getValue()).deepCopy()));
            });
        });
        return entries;
    }

    private static void validateShiftPreservation(ObjectNode root, ObjectNode sheet, RangeRef selection, int rowDelta, int columnDelta) {
        for (JsonNode sparkline : SnapshotMutationSupport.array(sheet, "sparklines")) validateBoundedPoint(sparkline.path("anchor").path("row").asInt(-1), sparkline.path("anchor").path("column").asInt(-1), selection, rowDelta, columnDelta, "sparkline");
        for (JsonNode spill : SnapshotMutationSupport.array(sheet, "spillRanges")) validateBoundedPoint(spill.path("anchor").path("row").asInt(-1), spill.path("anchor").path("column").asInt(-1), selection, rowDelta, columnDelta, "spill");
        for (JsonNode drawing : SnapshotMutationSupport.array(sheet, "drawings")) {
            JsonNode anchor = drawing.path("anchor");
            if ("absolute".equals(anchor.path("kind").asText())) continue;
            validateBoundedPoint(anchor.path("row").asInt(-1), anchor.path("column").asInt(-1), selection, rowDelta, columnDelta, "drawing");
            if (anchor.has("endRow") && anchor.has("endColumn")) validateBoundedPoint(anchor.path("endRow").asInt(-1), anchor.path("endColumn").asInt(-1), selection, rowDelta, columnDelta, "drawing extent");
        }
        for (JsonNode note : SnapshotMutationSupport.array(sheet, "notes")) validateBoundedPoint(note.path("row").asInt(-1), note.path("column").asInt(-1), selection, rowDelta, columnDelta, "note");
        for (JsonNode comment : SnapshotMutationSupport.array(sheet, "commentThreads")) validateBoundedPoint(comment.path("row").asInt(-1), comment.path("column").asInt(-1), selection, rowDelta, columnDelta, "comment");
        String sheetId = sheet.path("id").asText();
        for (JsonNode table : SnapshotMutationSupport.array(root, "tables")) {
            JsonNode range = table.get("sourceRange");
            if (range == null || !sheetId.equals(range.path("sheetId").asText())) continue;
            RangeRef ref = SnapshotMutationSupport.range(root, range);
            boolean contained = contains(selection, ref.startRow(), ref.startColumn()) && contains(selection, ref.endRow(), ref.endColumn());
            boolean intersects = intersects(selection, ref);
            if (intersects && !contained) throw ServiceException.validation("Cell shift partially intersects a workbook table");
            if (contained && (!contains(selection, ref.startRow() + rowDelta, ref.startColumn() + columnDelta) || !contains(selection, ref.endRow() + rowDelta, ref.endColumn() + columnDelta))) {
                throw ServiceException.validation("Cell shift would move a workbook table outside the selected range");
            }
        }
    }

    private static void validateBoundedPoint(int row, int column, RangeRef selection, int rowDelta, int columnDelta, String name) {
        if (!contains(selection, row, column)) return;
        if (!contains(selection, row + rowDelta, column + columnDelta)) throw ServiceException.validation("Cell shift would move " + name + " outside the selected range");
    }

    private static boolean intersects(RangeRef left, RangeRef right) {
        return left.sheetId().equals(right.sheetId())
                && left.startRow() <= right.endRow() && left.endRow() >= right.startRow()
                && left.startColumn() <= right.endColumn() && left.endColumn() >= right.startColumn();
    }

    private static void shiftBoundedMetadata(ObjectNode root, ObjectNode sheet, String sheetId, RangeRef selection, int rowDelta, int columnDelta) {
        for (JsonNode merge : SnapshotMutationSupport.array(sheet, "merges")) {
            ObjectNode object = requireObject(merge, "Merge");
            if (moveContainedRange(object.get("range"), selection, rowDelta, columnDelta)) {
                ObjectNode anchor = SnapshotMutationSupport.requiredObject(object, "anchor");
                anchor.put("row", anchor.path("row").asInt() + rowDelta);
                anchor.put("column", anchor.path("column").asInt() + columnDelta);
            }
        }
        for (String property : List.of("conditionalFormats", "dataValidations")) {
            for (JsonNode rule : SnapshotMutationSupport.array(sheet, property)) {
                for (JsonNode range : requireObject(rule, "Range rule").path("ranges")) moveContainedRange(range, selection, rowDelta, columnDelta);
            }
        }
        JsonNode filter = sheet.get("filter");
        if (filter != null && filter.isObject()) moveContainedRange(filter.get("range"), selection, rowDelta, columnDelta);
        for (JsonNode table : SnapshotMutationSupport.array(sheet, "sheetTables")) moveContainedRange(requireObject(table, "Sheet table").get("range"), selection, rowDelta, columnDelta);
        for (JsonNode table : SnapshotMutationSupport.array(root, "tables")) {
            JsonNode range = table.get("sourceRange");
            if (range != null && sheetId.equals(range.path("sheetId").asText())) moveContainedRange(range, selection, rowDelta, columnDelta);
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "pivots")) {
            ObjectNode pivot = requireObject(raw, "Pivot");
            SnapshotMutationSupport.validateKnownKeys(pivot, Set.of("schema", "id", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "nativeMetadata"), "Pivot");
            PivotMutationDescriptor.forEachWorksheetSourceRange(pivot, range -> moveContainedRange(range, selection, rowDelta, columnDelta));
            ObjectNode target = PivotMutationDescriptor.requiredTarget(pivot);
            JsonNode anchor = target.get("anchor");
            if (anchor != null && anchor.isObject() && contains(selection, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) {
                ((ObjectNode) anchor).put("row", anchor.path("row").asInt() + rowDelta).put("column", anchor.path("column").asInt() + columnDelta);
            }
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "sparklines")) {
            ObjectNode sparkline = requireObject(raw, "Sparkline");
            moveContainedRange(sparkline.get("sourceRange"), selection, rowDelta, columnDelta);
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(sparkline, "anchor");
            if (contains(selection, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) anchor.put("row", anchor.path("row").asInt() + rowDelta).put("column", anchor.path("column").asInt() + columnDelta);
        }
        for (JsonNode spill : SnapshotMutationSupport.array(sheet, "spillRanges")) {
            ObjectNode object = requireObject(spill, "Spill range");
            moveContainedRange(object.get("range"), selection, rowDelta, columnDelta);
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(object, "anchor");
            if (contains(selection, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) anchor.put("row", anchor.path("row").asInt() + rowDelta).put("column", anchor.path("column").asInt() + columnDelta);
        }
        for (JsonNode rule : SnapshotMutationSupport.array(sheet, "protectionRules")) {
            JsonNode range = rule.get("range");
            if (range != null) moveContainedRange(range, selection, rowDelta, columnDelta);
        }
        JsonNode banded = sheet.get("bandedRule");
        if (banded != null && banded.isObject()) moveContainedRange(banded.get("range"), selection, rowDelta, columnDelta);
        for (JsonNode drawing : SnapshotMutationSupport.array(sheet, "drawings")) {
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(requireObject(drawing, "Drawing"), "anchor");
            if ("absolute".equals(anchor.path("kind").asText()) || !contains(selection, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) continue;
            anchor.put("row", anchor.path("row").asInt() + rowDelta).put("column", anchor.path("column").asInt() + columnDelta);
            if (anchor.has("endRow")) anchor.put("endRow", anchor.path("endRow").asInt() + rowDelta);
            if (anchor.has("endColumn")) anchor.put("endColumn", anchor.path("endColumn").asInt() + columnDelta);
        }
        moveBoundedNotesAndComments(sheet, selection, rowDelta, columnDelta);
        if (rowDelta != 0) {
            remapBoundedIndices(sheet, "hiddenRows", "rowHeightsPx", selection.startRow(), selection.endRow(), rowDelta);
        }
        if (columnDelta != 0) {
            remapBoundedIndices(sheet, "hiddenColumns", "columnWidthsPx", selection.startColumn(), selection.endColumn(), columnDelta);
        }
    }

    private static boolean moveContainedRange(JsonNode raw, RangeRef selection, int rowDelta, int columnDelta) {
        if (raw == null || !raw.isObject()) return false;
        ObjectNode range = (ObjectNode) raw;
        if (!selection.sheetId().equals(range.path("sheetId").asText())) return false;
        if (!contains(selection, range.path("startRow").asInt(-1), range.path("startColumn").asInt(-1)) || !contains(selection, range.path("endRow").asInt(-1), range.path("endColumn").asInt(-1))) return false;
        range.put("startRow", range.path("startRow").asInt() + rowDelta);
        range.put("endRow", range.path("endRow").asInt() + rowDelta);
        range.put("startColumn", range.path("startColumn").asInt() + columnDelta);
        range.put("endColumn", range.path("endColumn").asInt() + columnDelta);
        return true;
    }

    private static void moveBoundedNotesAndComments(ObjectNode sheet, RangeRef selection, int rowDelta, int columnDelta) {
        for (int index = SnapshotMutationSupport.array(sheet, "notes").size() - 1; index >= 0; index--) {
            ObjectNode note = requireObject(SnapshotMutationSupport.array(sheet, "notes").get(index), "Note");
            int row = note.path("row").asInt(-1);
            int column = note.path("column").asInt(-1);
            if (!contains(selection, row, column)) continue;
            if (!contains(selection, row + rowDelta, column + columnDelta)) SnapshotMutationSupport.array(sheet, "notes").remove(index);
            else note.put("row", row + rowDelta).put("column", column + columnDelta);
        }
        for (int index = SnapshotMutationSupport.array(sheet, "commentThreads").size() - 1; index >= 0; index--) {
            ObjectNode thread = requireObject(SnapshotMutationSupport.array(sheet, "commentThreads").get(index), "Comment thread");
            int row = thread.path("row").asInt(-1);
            int column = thread.path("column").asInt(-1);
            if (!contains(selection, row, column)) continue;
            if (!contains(selection, row + rowDelta, column + columnDelta)) SnapshotMutationSupport.array(sheet, "commentThreads").remove(index);
            else thread.put("row", row + rowDelta).put("column", column + columnDelta);
        }
    }

    private static void remapBoundedIndices(ObjectNode sheet, String indexCollection, String sizeCollection, int start, int end, int delta) {
        ArrayNode indexes = SnapshotMutationSupport.array(sheet, indexCollection);
        ArrayNode nextIndexes = JsonNodeFactory.instance.arrayNode();
        for (JsonNode raw : indexes) {
            int value = raw.asInt(-1);
            if (value >= start && value <= end) value += delta;
            if (!containsNumber(nextIndexes, value)) nextIndexes.add(value);
        }
        sheet.set(indexCollection, nextIndexes);
        ObjectNode sizes = SnapshotMutationSupport.object(sheet, sizeCollection);
        ObjectNode nextSizes = JsonNodeFactory.instance.objectNode();
        sizes.fields().forEachRemaining(entry -> {
            int value = integerKey(entry.getKey(), SnapshotMutationSupport.MAX_ROW, sizeCollection);
            if (value >= start && value <= end) value += delta;
            nextSizes.set(Integer.toString(value), entry.getValue());
        });
        sheet.set(sizeCollection, nextSizes);
    }

    private static void rewriteMovedReferences(ObjectNode root, ObjectNode targetSheet, String targetSheetId, RangeRef selection, int rowDelta, int columnDelta) {
        FormulaReferenceTransformer.SheetIdentity target = identity(targetSheet);
        int destinationStartRow = selection.startRow() + rowDelta;
        int destinationEndRow = selection.endRow() + rowDelta;
        int destinationStartColumn = selection.startColumn() + columnDelta;
        int destinationEndColumn = selection.endColumn() + columnDelta;
        boolean destinationIsValid = destinationStartRow >= 0 && destinationEndRow <= SnapshotMutationSupport.MAX_ROW
                && destinationStartColumn >= 0 && destinationEndColumn <= SnapshotMutationSupport.MAX_COLUMN;
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode owner = requireObject(raw, "Sheet");
            FormulaReferenceTransformer.SheetIdentity ownerIdentity = identity(owner);
            ObjectNode cells = SnapshotMutationSupport.cells(owner);
            cells.fields().forEachRemaining(row -> ((ObjectNode) row.getValue()).fields().forEachRemaining(column -> {
                ObjectNode cell = requireObject(column.getValue(), "Cell");
                if (!cell.path("formula").isTextual()) return;
                int rowIndex = integerKey(row.getKey(), SnapshotMutationSupport.MAX_ROW, "Cell row");
                int columnIndex = integerKey(column.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                boolean insideDestination = destinationIsValid
                        && rowIndex >= destinationStartRow && rowIndex <= destinationEndRow
                        && columnIndex >= destinationStartColumn && columnIndex <= destinationEndColumn;
                if (owner.path("id").asText().equals(targetSheetId) && (contains(selection, rowIndex, columnIndex) || insideDestination)) return;
                String rewritten = FormulaReferenceTransformer.remapMovedRegion(cell.path("formula").asText(), ownerIdentity, target, new FormulaReferenceTransformer.Range(selection.startRow(), selection.endRow(), selection.startColumn(), selection.endColumn()), rowDelta, columnDelta);
                if (!rewritten.equals(cell.path("formula").asText())) {
                    cell.put("formula", rewritten);
                    cell.remove("formulaValue");
                }
            }));
        }
        ObjectNode names = SnapshotMutationSupport.object(root, "definedNames");
        names.fields().forEachRemaining(entry -> {
            if (entry.getValue().isTextual()) names.put(entry.getKey(), FormulaReferenceTransformer.remapMovedRegion(entry.getValue().asText(), target, target, new FormulaReferenceTransformer.Range(selection.startRow(), selection.endRow(), selection.startColumn(), selection.endColumn()), rowDelta, columnDelta));
        });
    }

    private static int[] validatePermutation(RangeRef range, ArrayNode sourceRows) {
        int[] rows = new int[sourceRows.size()];
        boolean[] seen = new boolean[rows.length];
        for (int index = 0; index < sourceRows.size(); index++) {
            JsonNode raw = sourceRows.get(index);
            if (!raw.isIntegralNumber() || !raw.canConvertToInt()) throw ServiceException.validation("Row permutation entry is invalid");
            int source = raw.intValue();
            if (source < range.startRow() || source > range.endRow()) throw ServiceException.validation("Row permutation source is outside range");
            int offset = source - range.startRow();
            if (seen[offset]) throw ServiceException.validation("Row permutation contains a duplicate source row");
            seen[offset] = true;
            rows[index] = source;
        }
        for (boolean value : seen) if (!value) throw ServiceException.validation("Row permutation must contain every row");
        return rows;
    }

    private static void validatePermutationPreservation(ObjectNode sheet, RangeRef range) {
        for (JsonNode merge : SnapshotMutationSupport.array(sheet, "merges")) {
            JsonNode mergeRange = merge.get("range");
            if (mergeRange != null && range.sheetId().equals(mergeRange.path("sheetId").asText())) {
                boolean intersects = mergeRange.path("startRow").asInt() <= range.endRow() && mergeRange.path("endRow").asInt() >= range.startRow()
                        && mergeRange.path("startColumn").asInt() <= range.endColumn() && mergeRange.path("endColumn").asInt() >= range.startColumn();
                if (intersects && !(mergeRange.path("startRow").asInt() >= range.startRow() && mergeRange.path("endRow").asInt() <= range.endRow())) {
                    throw ServiceException.validation("Row permutation partially intersects a merged range");
                }
            }
        }
        for (JsonNode table : SnapshotMutationSupport.array(sheet, "sheetTables")) {
            JsonNode tableRange = table.get("range");
            if (tableRange == null || !range.sheetId().equals(tableRange.path("sheetId").asText())) continue;
            boolean intersects = tableRange.path("startRow").asInt() <= range.endRow() && tableRange.path("endRow").asInt() >= range.startRow();
            if (intersects && !(tableRange.path("startRow").asInt() == range.startRow() && tableRange.path("endRow").asInt() == range.endRow())) {
                throw ServiceException.validation("Row permutation requires the complete sheet table range");
            }
        }
        JsonNode outline = sheet.get("outline");
        if (outline != null && outline.isObject()) {
            for (JsonNode group : ((ObjectNode) outline).path("groups")) {
                if (!"row".equals(group.path("axis").asText())) continue;
                boolean intersects = group.path("start").asInt() <= range.endRow() && group.path("end").asInt() >= range.startRow();
                if (intersects && !(group.path("start").asInt() >= range.startRow() && group.path("end").asInt() <= range.endRow())) {
                    throw ServiceException.validation("Row permutation partially intersects an outline group");
                }
            }
        }
    }

    private static void remapPermutedCells(ObjectNode sheet, RangeRef range, int[] sourceRows) {
        ObjectNode cells = SnapshotMutationSupport.cells(sheet);
        List<CellEntry> entries = new ArrayList<>();
        for (int row = range.startRow(); row <= range.endRow(); row++) {
            ObjectNode current = SnapshotMutationSupport.cellRow(cells, row, false);
            if (current == null) continue;
            int sourceRow = row;
            current.fields().forEachRemaining(column -> {
                int columnIndex = integerKey(column.getKey(), SnapshotMutationSupport.MAX_COLUMN, "Cell column");
                if (columnIndex >= range.startColumn() && columnIndex <= range.endColumn()) entries.add(new CellEntry(sourceRow, columnIndex, requireObject(column.getValue(), "Cell").deepCopy()));
            });
        }
        for (int row = range.startRow(); row <= range.endRow(); row++) {
            ObjectNode current = SnapshotMutationSupport.cellRow(cells, row, false);
            if (current == null) continue;
            for (int column = range.startColumn(); column <= range.endColumn(); column++) current.remove(Integer.toString(column));
            if (current.isEmpty()) cells.remove(Integer.toString(row));
        }
        for (int targetOffset = 0; targetOffset < sourceRows.length; targetOffset++) {
            int source = sourceRows[targetOffset];
            int target = range.startRow() + targetOffset;
            for (CellEntry entry : entries) if (entry.row() == source) SnapshotMutationSupport.putCell(sheet, new SnapshotMutationSupport.CellCoordinate(target, entry.column()), entry.cell());
        }
    }

    private static void remapPermutationMetadata(ObjectNode sheet, RangeRef range, int[] sourceRows) {
        int[] rowMap = new int[sourceRows.length];
        for (int targetOffset = 0; targetOffset < sourceRows.length; targetOffset++) rowMap[sourceRows[targetOffset] - range.startRow()] = range.startRow() + targetOffset;
        remapPermutationNotes(sheet, range, rowMap);
        remapPermutationComments(sheet, range, rowMap);
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "drawings")) remapDrawingRows(requireObject(raw, "Drawing"), range, rowMap);
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "sparklines")) {
            ObjectNode sparkline = requireObject(raw, "Sparkline");
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(sparkline, "anchor");
            anchor.put("row", remapRow(anchor.path("row").asInt(), range, rowMap));
            remapRangeRows(sparkline.get("sourceRange"), range, rowMap);
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "spillRanges")) {
            ObjectNode spill = requireObject(raw, "Spill range");
            SnapshotMutationSupport.requiredObject(spill, "anchor").put("row", remapRow(spill.path("anchor").path("row").asInt(), range, rowMap));
            remapRangeRows(spill.get("range"), range, rowMap);
        }
        for (String property : List.of("conditionalFormats", "dataValidations")) {
            for (JsonNode rule : SnapshotMutationSupport.array(sheet, property)) for (JsonNode ruleRange : requireObject(rule, "Range rule").path("ranges")) remapRangeRows(ruleRange, range, rowMap);
        }
        JsonNode filter = sheet.get("filter");
        if (filter != null && filter.isObject()) remapRangeRows(filter.get("range"), range, rowMap);
        for (JsonNode table : SnapshotMutationSupport.array(sheet, "sheetTables")) remapRangeRows(requireObject(table, "Sheet table").get("range"), range, rowMap);
        for (JsonNode rawPivot : SnapshotMutationSupport.array(sheet, "pivots")) {
            ObjectNode pivot = requireObject(rawPivot, "Pivot");
            SnapshotMutationSupport.validateKnownKeys(pivot, Set.of("schema", "id", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "nativeMetadata"), "Pivot");
            PivotMutationDescriptor.forEachWorksheetSourceRange(pivot, source -> remapRangeRows(source, range, rowMap));
            ObjectNode target = PivotMutationDescriptor.requiredTarget(pivot);
            ObjectNode anchor = PivotMutationDescriptor.requiredAnchor(target);
            if (range.sheetId().equals(target.path("sheetId").asText()) && contains(range, anchor.path("row").asInt(-1), anchor.path("column").asInt(-1))) {
                anchor.put("row", remapRow(anchor.path("row").asInt(), range, rowMap));
            }
        }
        for (JsonNode merge : SnapshotMutationSupport.array(sheet, "merges")) {
            ObjectNode object = requireObject(merge, "Merge");
            remapRangeRows(object.get("range"), range, rowMap);
            ObjectNode anchor = SnapshotMutationSupport.requiredObject(object, "anchor");
            anchor.put("row", remapRow(anchor.path("row").asInt(), range, rowMap));
        }
        JsonNode outline = sheet.get("outline");
        if (outline != null && outline.isObject()) {
            for (JsonNode group : ((ObjectNode) outline).path("groups")) {
                if (!"row".equals(group.path("axis").asText())) continue;
                int start = remapRow(group.path("start").asInt(), range, rowMap);
                int end = remapRow(group.path("end").asInt(), range, rowMap);
                ((ObjectNode) group).put("start", Math.min(start, end)).put("end", Math.max(start, end));
            }
        }
        for (JsonNode rule : SnapshotMutationSupport.array(sheet, "protectionRules")) if (rule.has("range")) remapRangeRows(rule.get("range"), range, rowMap);
    }

    private static void remapPermutationNotes(ObjectNode sheet, RangeRef range, int[] rowMap) {
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "notes")) {
            ObjectNode note = requireObject(raw, "Note");
            int row = note.path("row").asInt();
            int column = note.path("column").asInt();
            if (contains(range, row, column)) note.put("row", remapRow(row, range, rowMap));
        }
    }

    private static void remapPermutationComments(ObjectNode sheet, RangeRef range, int[] rowMap) {
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "commentThreads")) {
            ObjectNode comment = requireObject(raw, "Comment thread");
            int row = comment.path("row").asInt();
            if (row >= range.startRow() && row <= range.endRow()) comment.put("row", remapRow(row, range, rowMap));
        }
    }

    private static void remapDrawingRows(ObjectNode drawing, RangeRef range, int[] rowMap) {
        ObjectNode anchor = SnapshotMutationSupport.requiredObject(drawing, "anchor");
        if ("absolute".equals(anchor.path("kind").asText())) return;
        if (anchor.has("row")) anchor.put("row", remapRow(anchor.path("row").asInt(), range, rowMap));
        if (anchor.has("endRow")) anchor.put("endRow", remapRow(anchor.path("endRow").asInt(), range, rowMap));
    }

    private static void remapRangeRows(JsonNode raw, RangeRef range, int[] rowMap) {
        if (raw == null || !raw.isObject() || !range.sheetId().equals(raw.path("sheetId").asText())) return;
        int start = raw.path("startRow").asInt();
        int end = raw.path("endRow").asInt();
        if (start < range.startRow() || start > range.endRow() || end < range.startRow() || end > range.endRow()) return;
        int remappedStart = remapRow(start, range, rowMap);
        int remappedEnd = remapRow(end, range, rowMap);
        ((ObjectNode) raw).put("startRow", Math.min(remappedStart, remappedEnd)).put("endRow", Math.max(remappedStart, remappedEnd));
    }

    private static int remapRow(int row, RangeRef range, int[] rowMap) {
        if (row < range.startRow() || row > range.endRow()) return row;
        return rowMap[row - range.startRow()];
    }

    private static ObjectNode requireObject(JsonNode value, String label) {
        if (value == null || !value.isObject()) throw ServiceException.validation(label + " must be an object");
        return (ObjectNode) value;
    }

    enum ShiftDirection {
        DOWN, UP, RIGHT, LEFT;

        ShiftDirection reverse() {
            return switch (this) {
                case DOWN -> UP;
                case UP -> DOWN;
                case RIGHT -> LEFT;
                case LEFT -> RIGHT;
            };
        }
    }

    private record CellEntry(int row, int column, ObjectNode cell) {
    }
}
