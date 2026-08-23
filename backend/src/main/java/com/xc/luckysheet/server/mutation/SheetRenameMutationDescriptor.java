package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.List;

/** Renames a sheet and rewrites only parsed qualified formula references. */
final class SheetRenameMutationDescriptor extends CanonicalJsonMutationDescriptor {
    SheetRenameMutationDescriptor() {
        super("sheet.rename", WorkbookAclRole.EDITOR, true, "structure");
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        SnapshotMutationSupport.params(mutation);
        return List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        ObjectNode target = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        String oldName = SnapshotMutationSupport.text(target, "name");
        String newName = SnapshotMutationSupport.text(params, "name").trim();
        if (newName.isBlank() || newName.length() > 31) throw ServiceException.validation("Worksheet name is invalid");
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            if (raw.isObject() && !mutation.sheetId().equals(raw.path("id").asText()) && newName.equalsIgnoreCase(raw.path("name").asText())) {
                throw ServiceException.conflict("Worksheet name already exists: " + newName);
            }
        }
        if (oldName.equals(newName)) return root;
        target.put("name", newName);
        rewriteWorkbookFormulaReferences(root, oldName, newName);
        return root;
    }

    private void rewriteWorkbookFormulaReferences(ObjectNode root, String oldName, String newName) {
        for (JsonNode raw : SnapshotMutationSupport.sheets(root)) {
            ObjectNode sheet = requireObject(raw, "Sheet");
            ObjectNode cells = SnapshotMutationSupport.cells(sheet);
            cells.fields().forEachRemaining(row -> {
                ObjectNode cellRow = requireObject(row.getValue(), "Cell row");
                cellRow.fields().forEachRemaining(column -> {
                    ObjectNode cell = requireObject(column.getValue(), "Cell");
                    if (!cell.path("formula").isTextual()) return;
                    String original = cell.path("formula").asText();
                    String rewritten = FormulaReferenceTransformer.renameSheet(original, oldName, newName);
                    if (!original.equals(rewritten)) {
                        cell.put("formula", rewritten);
                        cell.remove("formulaValue");
                    }
                });
            });
        }
        ObjectNode names = SnapshotMutationSupport.object(root, "definedNames");
        names.fields().forEachRemaining(entry -> {
            if (entry.getValue().isTextual()) names.put(entry.getKey(), FormulaReferenceTransformer.renameSheet(entry.getValue().asText(), oldName, newName));
        });
        for (JsonNode raw : SnapshotMutationSupport.array(root, "definedNameModels")) {
            ObjectNode model = requireObject(raw, "Defined name");
            if (model.path("formula").isTextual()) model.put("formula", FormulaReferenceTransformer.renameSheet(model.path("formula").asText(), oldName, newName));
        }
    }

    private ObjectNode requireObject(JsonNode value, String label) {
        if (value == null || !value.isObject()) throw ServiceException.validation(label + " must be an object");
        return (ObjectNode) value;
    }
}
