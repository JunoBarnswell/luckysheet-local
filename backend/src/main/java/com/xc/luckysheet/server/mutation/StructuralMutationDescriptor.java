package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.List;
import java.util.Set;

/** Server reducers for whole-axis and canonical cell-band structural worksheet mutations. */
final class StructuralMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted",
            "cells.inserted", "cells.deleted", "cells.inserted.restore", "cells.deleted.restore", "rows.permuted"
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
            case "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted" -> List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
            case "cells.inserted", "cells.deleted" -> List.of(cellAffectedBand(root, mutation.sheetId(), params));
            case "cells.inserted.restore", "cells.deleted.restore" -> List.of(restoreAffectedBand(root, mutation.sheetId(), params));
            case "rows.permuted" -> {
                RangeRef selected = ownRange(root, mutation.sheetId(), params);
                // Row permutation remaps row-addressed metadata such as
                // protection rules across the full Excel coordinate domain,
                // not merely the materialized worksheet width.
                yield List.of(new RangeRef(selected.sheetId(), selected.startRow(), selected.endRow(), 0, SnapshotMutationSupport.MAX_COLUMN));
            }
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
            case "cells.inserted", "cells.deleted" -> applyCellShift(root, mutation.sheetId(), params);
            case "cells.inserted.restore", "cells.deleted.restore" -> restore(root, mutation.sheetId(), params);
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
        JsonNode spec = params.get("spec");
        StructuralSnapshotReducer.restoreShiftedCells(root, sheetId, spec, params.get("cells"));
    }

    private void applyCellShift(ObjectNode root, String sheetId, ObjectNode params) {
        RangeRef range = ownRange(root, sheetId, params);
        RangeRef band = ownRangeField(root, sheetId, params, "affectedBand");
        String operation = text(params.get("operation"), "Cell shift operation");
        String axis = text(params.get("axis"), "Cell shift axis");
        StructuralSnapshotReducer.shiftCells(root, sheetId, range, operation, axis, band);
    }

    private RangeRef cellAffectedBand(ObjectNode root, String sheetId, ObjectNode params) {
        return ownRangeField(root, sheetId, params, "affectedBand");
    }

    private RangeRef restoreAffectedBand(ObjectNode root, String sheetId, ObjectNode params) {
        JsonNode spec = params.get("spec");
        if (spec == null || !spec.isObject()) throw ServiceException.validation("Structural restore spec is required");
        return ownRangeField(root, sheetId, (ObjectNode) spec, "affectedBand");
    }

    private RangeRef ownRange(ObjectNode root, String sheetId, ObjectNode params) {
        RangeRef range = SnapshotMutationSupport.range(root, params.get("range"));
        SnapshotMutationSupport.requireSheet(range, sheetId);
        return range;
    }

    private RangeRef ownRangeField(ObjectNode root, String sheetId, ObjectNode params, String field) {
        RangeRef range = SnapshotMutationSupport.range(root, params.get(field));
        SnapshotMutationSupport.requireSheet(range, sheetId);
        return range;
    }

    private String text(JsonNode value, String label) {
        if (value == null || !value.isTextual() || value.asText().isBlank()) throw ServiceException.validation(label + " is invalid");
        return value.asText();
    }

    private int integer(JsonNode value, String label) {
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 0) throw ServiceException.validation(label + " is invalid");
        return value.intValue();
    }
}
