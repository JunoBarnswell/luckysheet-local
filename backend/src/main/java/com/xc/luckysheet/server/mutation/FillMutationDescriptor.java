package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

/**
 * Server reducer for the canonical fill command.
 *
 * The browser sends the side-effect-free planner output as before/after cell
 * snapshots. The server derives the affected target band, validates the
 * one-axis geometry and verifies every before image while the workbook
 * operation lock is held. A stale or widened payload therefore fails before
 * the snapshot returned by the reducer can be committed.
 */
public final class FillMutationDescriptor extends CanonicalJsonMutationDescriptor {
    public static final Set<String> IDS = Set.of("fill.applied", "fill.restored");

    public FillMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, true, "edit-cell");
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unknown fill mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        return List.of(validatePayload(root, mutation.sheetId(), params));
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        RangeRef target = validatePayload(root, mutation.sheetId(), params);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        ArrayNode writes = (ArrayNode) params.get("writes");

        // Validate all snapshots before applying the first write. The reducer
        // itself is pure over a copied snapshot, but this also makes the
        // fail-close boundary explicit for a malformed mixed payload.
        for (JsonNode value : writes) {
            ObjectNode entry = requireWrite(value);
            SnapshotMutationSupport.CellCoordinate coordinate = writeCoordinate(root, mutation.sheetId(), target, entry);
            JsonNode current = SnapshotMutationSupport.cell(sheet, coordinate, false);
            assertBefore(current, entry.get("before"), coordinate);
            assertAfter(entry.get("after"));
        }
        for (JsonNode value : writes) {
            ObjectNode entry = (ObjectNode) value;
            SnapshotMutationSupport.CellCoordinate coordinate = writeCoordinate(root, mutation.sheetId(), target, entry);
            JsonNode after = entry.get("after");
            if (after == null || after.isNull()) SnapshotMutationSupport.removeCell(sheet, coordinate);
            else SnapshotMutationSupport.putCell(sheet, coordinate, after);
        }
        return root;
    }

    private RangeRef validatePayload(ObjectNode root, String sheetId, ObjectNode params) {
        if (!sheetId.equals(params.path("sheetId").asText())) throw ServiceException.validation("Fill sheetId does not match mutation sheetId");
        RangeRef source = requireOwnRange(root, sheetId, params, "sourceRange");
        RangeRef target = requireOwnRange(root, sheetId, params, "targetRange");
        String direction = SnapshotMutationSupport.text(params, "direction");
        String mode = SnapshotMutationSupport.text(params, "mode");
        if (!Set.of("down", "up", "right", "left").contains(direction)) throw ServiceException.validation("Fill direction is invalid");
        if (!Set.of("copy", "series").contains(mode)) throw ServiceException.validation("Fill mode is invalid");
        assertGeometry(source, target, direction);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        int rowCount = sheet.path("rowCount").asInt(-1);
        int columnCount = sheet.path("columnCount").asInt(-1);
        if (rowCount < 1 || columnCount < 1 || target.endRow() >= rowCount || target.endColumn() >= columnCount) {
            throw ServiceException.validation("Fill target is outside worksheet bounds");
        }
        JsonNode writesNode = params.get("writes");
        if (writesNode == null || !writesNode.isArray() || writesNode.isEmpty() || writesNode.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) {
            throw ServiceException.validation("Fill writes are required and bounded");
        }
        Set<String> coordinates = new HashSet<>();
        for (JsonNode value : writesNode) {
            ObjectNode entry = requireWrite(value);
            SnapshotMutationSupport.CellCoordinate coordinate = writeCoordinate(root, sheetId, target, entry);
            if (!coordinates.add(coordinate.row() + ":" + coordinate.column())) throw ServiceException.validation("Fill writes contain a duplicate coordinate");
            assertAfter(entry.get("after"));
        }
        return target;
    }

    private RangeRef requireOwnRange(ObjectNode root, String sheetId, ObjectNode params, String property) {
        RangeRef range = SnapshotMutationSupport.range(root, params.get(property));
        SnapshotMutationSupport.requireSheet(range, sheetId);
        if (SnapshotMutationSupport.cellCount(range) > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Fill range is too large");
        return range;
    }

    private void assertGeometry(RangeRef source, RangeRef target, String direction) {
        boolean contains = target.startRow() <= source.startRow() && target.endRow() >= source.endRow()
                && target.startColumn() <= source.startColumn() && target.endColumn() >= source.endColumn();
        if (!contains) throw ServiceException.validation("Fill target must contain source range");
        boolean sameColumns = source.startColumn() == target.startColumn() && source.endColumn() == target.endColumn();
        boolean sameRows = source.startRow() == target.startRow() && source.endRow() == target.endRow();
        boolean valid = switch (direction) {
            case "down" -> sameColumns && target.startRow() == source.startRow() && target.endRow() >= source.endRow();
            case "up" -> sameColumns && target.endRow() == source.endRow() && target.startRow() <= source.startRow();
            case "right" -> sameRows && target.startColumn() == source.startColumn() && target.endColumn() >= source.endColumn();
            case "left" -> sameRows && target.endColumn() == source.endColumn() && target.startColumn() <= source.startColumn();
            default -> false;
        };
        if (!valid) throw ServiceException.validation("Fill direction requires a contiguous one-axis target extension");
    }

    private ObjectNode requireWrite(JsonNode value) {
        if (value == null || !value.isObject()) throw ServiceException.validation("Fill write must be an object");
        ObjectNode entry = (ObjectNode) value;
        if (!entry.has("row") || !entry.path("row").isIntegralNumber()
                || !entry.has("column") || !entry.path("column").isIntegralNumber()) {
            throw ServiceException.validation("Fill write coordinates are required");
        }
        return entry;
    }

    private SnapshotMutationSupport.CellCoordinate writeCoordinate(
            ObjectNode root,
            String sheetId,
            RangeRef target,
            ObjectNode entry
    ) {
        int row = SnapshotMutationSupport.index(root, sheetId, entry, "row");
        int column = SnapshotMutationSupport.index(root, sheetId, entry, "column");
        SnapshotMutationSupport.CellCoordinate coordinate = new SnapshotMutationSupport.CellCoordinate(row, column);
        if (!SnapshotMutationSupport.contains(target, coordinate)) throw ServiceException.validation("Fill write is outside target range");
        return coordinate;
    }

    private void assertBefore(JsonNode current, JsonNode before, SnapshotMutationSupport.CellCoordinate coordinate) {
        if (before == null || before.isNull()) {
            if (current != null) throw ServiceException.conflict("Fill target changed at " + coordinate.row() + ":" + coordinate.column());
            return;
        }
        if (!before.isObject() || current == null || !current.equals(before)) {
            throw ServiceException.conflict("Fill target changed at " + coordinate.row() + ":" + coordinate.column());
        }
    }

    private void assertAfter(JsonNode after) {
        if (after == null || after.isNull()) return;
        if (!after.isObject() || !after.has("value")) throw ServiceException.validation("Fill after cell must contain a value");
    }
}
