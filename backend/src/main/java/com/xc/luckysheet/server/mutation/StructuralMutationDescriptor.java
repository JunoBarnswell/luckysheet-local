package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.List;
import java.util.Set;

/** Server reducers for whole-axis and bounded structural worksheet mutations. */
final class StructuralMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted",
            "cells.shifted", "cells.shifted.restore", "rows.permuted"
    );

    StructuralMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, true, "structure");
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported structural mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        return switch (id()) {
            case "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted", "cells.shifted", "cells.shifted.restore" -> List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
            case "rows.permuted" -> List.of(ownRange(root, mutation.sheetId(), params));
            default -> throw ServiceException.validation("Unsupported structural mutation: " + id());
        };
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        switch (id()) {
            case "rows.inserted" -> axis(root, mutation.sheetId(), params, FormulaReferenceTransformer.Axis.ROW, FormulaReferenceTransformer.Direction.INSERT);
            case "rows.deleted" -> axis(root, mutation.sheetId(), params, FormulaReferenceTransformer.Axis.ROW, FormulaReferenceTransformer.Direction.DELETE);
            case "columns.inserted" -> axis(root, mutation.sheetId(), params, FormulaReferenceTransformer.Axis.COLUMN, FormulaReferenceTransformer.Direction.INSERT);
            case "columns.deleted" -> axis(root, mutation.sheetId(), params, FormulaReferenceTransformer.Axis.COLUMN, FormulaReferenceTransformer.Direction.DELETE);
            case "cells.shifted" -> StructuralSnapshotReducer.shiftCells(root, mutation.sheetId(), ownRange(root, mutation.sheetId(), params), direction(params));
            case "cells.shifted.restore" -> restore(root, mutation.sheetId(), params);
            case "rows.permuted" -> StructuralSnapshotReducer.permuteRows(root, mutation.sheetId(), ownRange(root, mutation.sheetId(), params), params.get("sourceRows"));
            default -> throw ServiceException.validation("Unsupported structural mutation: " + id());
        }
        return root;
    }

    private void axis(ObjectNode root, String sheetId, ObjectNode params, FormulaReferenceTransformer.Axis axis, FormulaReferenceTransformer.Direction direction) {
        int at = integer(params.get("at"), "Structural at");
        int count = integer(params.get("count"), "Structural count");
        if (count < 1) throw ServiceException.validation("Structural count must be positive");
        StructuralSnapshotReducer.applyAxis(root, sheetId, axis, at, count, direction);
    }

    private void restore(ObjectNode root, String sheetId, ObjectNode params) {
        RangeRef range = ownRange(root, sheetId, params);
        JsonNode direction = params.get("direction");
        StructuralSnapshotReducer.ShiftDirection original = direction == null || direction.isNull() ? null : direction(direction);
        StructuralSnapshotReducer.restoreShiftedCells(root, sheetId, range, original, params.get("cells"));
    }

    private StructuralSnapshotReducer.ShiftDirection direction(ObjectNode params) {
        JsonNode value = params.get("direction");
        if (value == null) throw ServiceException.validation("Shift direction is required");
        return direction(value);
    }

    private StructuralSnapshotReducer.ShiftDirection direction(JsonNode value) {
        if (!value.isTextual()) throw ServiceException.validation("Shift direction is invalid");
        return switch (value.asText()) {
            case "down" -> StructuralSnapshotReducer.ShiftDirection.DOWN;
            case "up" -> StructuralSnapshotReducer.ShiftDirection.UP;
            case "right" -> StructuralSnapshotReducer.ShiftDirection.RIGHT;
            case "left" -> StructuralSnapshotReducer.ShiftDirection.LEFT;
            default -> throw ServiceException.validation("Shift direction is invalid");
        };
    }

    private RangeRef ownRange(ObjectNode root, String sheetId, ObjectNode params) {
        RangeRef range = SnapshotMutationSupport.range(root, params.get("range"));
        SnapshotMutationSupport.requireSheet(range, sheetId);
        return range;
    }

    private int integer(JsonNode value, String label) {
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 0) throw ServiceException.validation(label + " is invalid");
        return value.intValue();
    }
}
