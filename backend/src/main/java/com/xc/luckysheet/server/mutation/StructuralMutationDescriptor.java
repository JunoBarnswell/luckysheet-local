package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.contract.DataRegionContextValidator;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.List;
import java.util.Set;

/** Server reducers for whole-axis and canonical cell-band structural worksheet mutations. */
final class StructuralMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted",
            "cells.inserted", "cells.deleted", "cells.inserted.restore", "cells.deleted.restore", "rows.permuted", "range.move"
    );

    StructuralMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR);
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported structural mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        return switch (id()) {
            case "rows.inserted", "rows.deleted" -> List.of(axisAffectedRange(
                    root, mutation.sheetId(), params, FormulaReferenceTransformer.Axis.ROW, id().equals("rows.inserted")));
            case "columns.inserted", "columns.deleted" -> List.of(axisAffectedRange(
                    root, mutation.sheetId(), params, FormulaReferenceTransformer.Axis.COLUMN, id().equals("columns.inserted")));
            case "cells.inserted", "cells.deleted" -> List.of(cellAffectedBand(root, mutation.sheetId(), params));
            case "cells.inserted.restore", "cells.deleted.restore" -> List.of(restoreAffectedBand(root, mutation.sheetId(), params));
            case "rows.permuted" -> {
                DataRegionContextValidator.validateSort(root, mutation.sheetId(), params);
                RangeRef selected = ownRange(root, mutation.sheetId(), params);
                int declaredEndColumn = integer(params.get("affectedColumnEnd"), "Rows permutation affected column end");
                int canonicalEndColumn = SheetRuleLifecycle.affectedColumnEnd(root, SnapshotMutationSupport.sheet(root, mutation.sheetId()), selected.endColumn(), selected.startRow(), selected.endRow());
                if (declaredEndColumn < canonicalEndColumn || declaredEndColumn > SnapshotMutationSupport.MAX_COLUMN) {
                    throw ServiceException.validation("Rows permutation affected column extent does not cover current worksheet metadata");
                }
                yield List.of(new RangeRef(selected.sheetId(), selected.startRow(), selected.endRow(), 0, declaredEndColumn));
            }
            case "range.move" -> {
                RangeRef source = ownRangeField(root, mutation.sheetId(), params, "sourceRange");
                RangeRef target = moveTarget(root, mutation.sheetId(), source, params.get("targetOrigin"));
                yield List.of(source, target);
            }
            default -> throw ServiceException.validation("Unsupported structural mutation: " + id());
        };
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        return applyWithPatch(snapshot, mutation).snapshot();
    }

    @Override
    public MutationApplication applyWithPatch(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        StructuralPatch structuralPatch = null;
        switch (id()) {
            case "rows.inserted" -> structuralPatch = axis(root, mutation.sheetId(), mutation.id(), params, FormulaReferenceTransformer.Axis.ROW, FormulaReferenceTransformer.Direction.INSERT);
            case "rows.deleted" -> structuralPatch = axis(root, mutation.sheetId(), mutation.id(), params, FormulaReferenceTransformer.Axis.ROW, FormulaReferenceTransformer.Direction.DELETE);
            case "columns.inserted" -> structuralPatch = axis(root, mutation.sheetId(), mutation.id(), params, FormulaReferenceTransformer.Axis.COLUMN, FormulaReferenceTransformer.Direction.INSERT);
            case "columns.deleted" -> structuralPatch = axis(root, mutation.sheetId(), mutation.id(), params, FormulaReferenceTransformer.Axis.COLUMN, FormulaReferenceTransformer.Direction.DELETE);
            case "cells.inserted", "cells.deleted" -> structuralPatch = applyCellShift(root, mutation.sheetId(), mutation.id(), params);
            case "cells.inserted.restore", "cells.deleted.restore" -> restore(root, mutation.sheetId(), mutation.id(), params);
            case "rows.permuted" -> {
                DataRegionContextValidator.validateSort(root, mutation.sheetId(), params);
                RangeRef selected = ownRange(root, mutation.sheetId(), params);
                int declaredEndColumn = integer(params.get("affectedColumnEnd"), "Rows permutation affected column end");
                int canonicalEndColumn = SheetRuleLifecycle.affectedColumnEnd(root, SnapshotMutationSupport.sheet(root, mutation.sheetId()), selected.endColumn(), selected.startRow(), selected.endRow());
                if (declaredEndColumn < canonicalEndColumn || declaredEndColumn > SnapshotMutationSupport.MAX_COLUMN) {
                    throw ServiceException.validation("Rows permutation affected column extent does not cover current worksheet metadata");
                }
                StructuralSnapshotReducer.permuteRows(root, mutation.sheetId(), selected, declaredEndColumn, params.get("sourceRows"));
            }
            case "range.move" -> {
                SnapshotMutationSupport.validateKnownKeys(params, Set.of("sheetId", "sourceRange", "targetOrigin"), "range.move");
                RangeRef source = ownRangeField(root, mutation.sheetId(), params, "sourceRange");
                RangeRef target = moveTarget(root, mutation.sheetId(), source, params.get("targetOrigin"));
                StructuralSnapshotReducer.moveRange(root, mutation.sheetId(), source, target);
            }
            default -> throw ServiceException.validation("Unsupported structural mutation: " + id());
        }
        return new MutationApplication(root, structuralPatch);
    }

    private StructuralPatch axis(ObjectNode root, String sheetId, String mutationId, ObjectNode params, FormulaReferenceTransformer.Axis axis, FormulaReferenceTransformer.Direction direction) {
        int at = integer(params.get("at"), "Structural at");
        int count = integer(params.get("count"), "Structural count");
        if (count < 1) throw ServiceException.validation("Structural count must be positive");
        return StructuralSnapshotReducer.applyAxis(root, sheetId, mutationId, axis, at, count, direction);
    }

    private RangeRef axisAffectedRange(
            ObjectNode root,
            String sheetId,
            ObjectNode params,
            FormulaReferenceTransformer.Axis axis,
            boolean insert
    ) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        int at = integer(params.get("at"), "Structural at");
        int count = integer(params.get("count"), "Structural count");
        if (at < 0 || count < 1) throw ServiceException.validation("Structural axis range is invalid");

        boolean rows = axis == FormulaReferenceTransformer.Axis.ROW;
        int dimension = SnapshotMutationSupport.canonicalDimension(sheet, rows ? "rowCount" : "columnCount");
        int orthogonalEnd = SnapshotMutationSupport.canonicalDimension(sheet, rows ? "columnCount" : "rowCount") - 1;
        int maximumIndex = rows ? SnapshotMutationSupport.MAX_ROW : SnapshotMutationSupport.MAX_COLUMN;
        long end = (long) at + count - 1;
        boolean valid = insert
                ? at <= dimension && (long) dimension + count <= (long) maximumIndex + 1
                : at < dimension && count <= dimension - at;
        if (!valid || end > maximumIndex) throw ServiceException.validation("Structural axis range exceeds worksheet bounds");

        return rows
                ? new RangeRef(sheetId, at, (int) end, 0, orthogonalEnd)
                : new RangeRef(sheetId, 0, orthogonalEnd, at, (int) end);
    }

    private void restore(ObjectNode root, String sheetId, String mutationId, ObjectNode params) {
        JsonNode spec = params.get("spec");
        StructuralSnapshotReducer.restoreShiftedCells(root, sheetId, mutationId, spec, params.get("cells"));
    }

    private StructuralPatch applyCellShift(ObjectNode root, String sheetId, String mutationId, ObjectNode params) {
        RangeRef range = ownRange(root, sheetId, params);
        RangeRef band = ownRangeField(root, sheetId, params, "affectedBand");
        String operation = text(params.get("operation"), "Cell shift operation");
        String axis = text(params.get("axis"), "Cell shift axis");
        return StructuralSnapshotReducer.shiftCells(root, sheetId, mutationId, range, operation, axis, band);
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

    private RangeRef moveTarget(ObjectNode root, String sheetId, RangeRef source, JsonNode rawOrigin) {
        if (rawOrigin == null || !rawOrigin.isObject()) throw ServiceException.validation("Move target origin must be an object");
        ObjectNode origin = (ObjectNode) rawOrigin;
        SnapshotMutationSupport.validateKnownKeys(origin, Set.of("row", "column"), "Move target origin");
        int row = integer(origin.get("row"), "Move target row");
        int column = integer(origin.get("column"), "Move target column");
        long endRow = (long) row + source.endRow() - source.startRow();
        long endColumn = (long) column + source.endColumn() - source.startColumn();
        if (endRow > SnapshotMutationSupport.MAX_ROW || endColumn > SnapshotMutationSupport.MAX_COLUMN) {
            throw ServiceException.validation("Move target exceeds worksheet bounds");
        }
        SnapshotMutationSupport.sheet(root, sheetId);
        return new RangeRef(sheetId, row, (int) endRow, column, (int) endColumn);
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
