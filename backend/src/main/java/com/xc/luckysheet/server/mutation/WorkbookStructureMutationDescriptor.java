package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.List;
import java.util.Set;

/** Canonical snapshot reducers for worksheet lifecycle and persisted hyperlinks. */
final class WorkbookStructureMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "sheet.add", "sheet.remove", "sheet.rename", "sheet.duplicated", "sheet.restore",
            "hyperlink.set", "hyperlink.remove"
    );

    WorkbookStructureMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, id.startsWith("hyperlink."), id.startsWith("hyperlink.") ? "edit-cell" : "structure");
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        return switch (id()) {
            case "hyperlink.set", "hyperlink.remove" -> List.of(SnapshotMutationSupport.cellRange(root, mutation.sheetId(), params));
            default -> List.of();
        };
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        switch (id()) {
            case "sheet.add" -> add(root, mutation.sheetId(), params);
            case "sheet.remove" -> remove(root, params);
            case "sheet.rename" -> rename(root, mutation.sheetId(), params);
            case "sheet.duplicated" -> duplicate(root, params);
            case "sheet.restore" -> restore(root, params);
            case "hyperlink.set" -> setHyperlink(root, mutation.sheetId(), params);
            case "hyperlink.remove" -> removeHyperlink(root, mutation.sheetId(), params);
            default -> throw ServiceException.validation("Unsupported workbook structure mutation: " + id());
        }
        return root;
    }

    private void add(ObjectNode root, String mutationSheetId, ObjectNode params) {
        String id = SnapshotMutationSupport.text(params, "id");
        String name = SnapshotMutationSupport.text(params, "name").trim();
        if (!mutationSheetId.equals(id) || name.isBlank()) throw ServiceException.validation("sheet.add identity is invalid");
        if (findSheetIndex(root, id) >= 0) throw ServiceException.conflict("Sheet already exists: " + id);
        int rows = params.has("rowCount") ? dimension(params.get("rowCount"), "rowCount") : 1000;
        int columns = params.has("columnCount") ? dimension(params.get("columnCount"), "columnCount") : 26;
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        ObjectNode sheet = JsonNodeFactory.instance.objectNode();
        sheet.put("kind", "worksheet");
        sheet.put("id", id);
        sheet.put("name", name);
        sheet.put("rowCount", rows);
        sheet.put("columnCount", columns);
        sheet.putObject("cells");
        sheet.putArray("dataRegions");
        sheet.putArray("merges");
        sheet.putObject("pane").put("kind", "none");
        sheet.put("defaultRowHeightPx", 20);
        sheet.put("defaultColumnWidthPx", 64);
        sheet.putObject("rowHeightsPx");
        sheet.putObject("columnWidthsPx");
        sheet.putArray("pivots");
        sheet.putArray("sparklines");
        sheet.putArray("drawings");
        sheet.putObject("drawingPayloads");
        sheet.putArray("hyperlinks");
        sheet.putArray("notes");
        sheet.putArray("commentThreads");
        sheets.add(sheet);
    }

    private void remove(ObjectNode root, ObjectNode params) {
        String id = SnapshotMutationSupport.text(params, "id");
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        if (sheets.size() <= 1) throw ServiceException.validation("A workbook must keep at least one worksheet");
        int index = findSheetIndex(root, id);
        if (index < 0) throw ServiceException.notFound("Sheet not found: " + id);
        sheets.remove(index);
        removeSheetScopedDocuments(root, id);
    }

    private void rename(ObjectNode root, String mutationSheetId, ObjectNode params) {
        String sheetId = SnapshotMutationSupport.text(params, "sheetId");
        String name = SnapshotMutationSupport.text(params, "name").trim();
        if (!mutationSheetId.equals(sheetId) || name.isBlank()) throw ServiceException.validation("sheet.rename identity is invalid");
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        String previousName = sheet.path("name").asText();
        if (!previousName.equals(name)) rewriteSheetReferences(root, previousName, name);
        sheet.put("name", name);
    }

    private void duplicate(ObjectNode root, ObjectNode params) {
        String sourceSheetId = SnapshotMutationSupport.text(params, "sourceSheetId");
        String newId = SnapshotMutationSupport.text(params, "newId");
        String newName = SnapshotMutationSupport.text(params, "newName").trim();
        if (newName.isBlank() || findSheetIndex(root, newId) >= 0) throw ServiceException.conflict("Duplicate sheet identity is invalid");
        int sourceIndex = findSheetIndex(root, sourceSheetId);
        if (sourceIndex < 0) throw ServiceException.notFound("Sheet not found: " + sourceSheetId);
        ObjectNode copy = ((ObjectNode) SnapshotMutationSupport.sheets(root).get(sourceIndex)).deepCopy();
        replaceSheetReferences(copy, sourceSheetId, newId);
        copy.put("id", newId);
        copy.put("name", newName);
        SnapshotMutationSupport.sheets(root).insert(sourceIndex + 1, copy);
    }

    private void restore(ObjectNode root, ObjectNode params) {
        JsonNode raw = params.get("sheet");
        if (raw == null || !raw.isObject()) throw ServiceException.validation("sheet.restore requires a canonical sheet snapshot");
        ObjectNode sheet = ((ObjectNode) raw).deepCopy();
        String id = sheet.path("id").asText().trim();
        if (id.isBlank() || findSheetIndex(root, id) >= 0) throw ServiceException.conflict("Restored sheet identity is invalid");
        requireSheetShape(sheet);
        int index = params.path("index").isInt() ? params.path("index").intValue() : SnapshotMutationSupport.sheets(root).size();
        int bounded = Math.max(0, Math.min(index, SnapshotMutationSupport.sheets(root).size()));
        SnapshotMutationSupport.sheets(root).insert(bounded, sheet);
    }

    private void setHyperlink(ObjectNode root, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
        JsonNode hyperlink = params.get("hyperlink");
        if (hyperlink == null || !hyperlink.isObject() || hyperlink.path("id").asText().isBlank()
                || !hyperlink.path("target").isObject()) throw ServiceException.validation("hyperlink.set requires a canonical hyperlink");
        ArrayNode hyperlinks = SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, sheetId), "hyperlinks");
        removeHyperlinkAt(hyperlinks, coordinate);
        ObjectNode entry = hyperlinks.objectNode();
        entry.put("row", coordinate.row());
        entry.put("column", coordinate.column());
        entry.set("hyperlink", hyperlink.deepCopy());
        hyperlinks.add(entry);
    }

    private void removeHyperlink(ObjectNode root, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
        ArrayNode hyperlinks = SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, sheetId), "hyperlinks");
        if (!removeHyperlinkAt(hyperlinks, coordinate)) throw ServiceException.notFound("Hyperlink not found");
    }

    private boolean removeHyperlinkAt(ArrayNode hyperlinks, SnapshotMutationSupport.CellCoordinate coordinate) {
        for (int index = 0; index < hyperlinks.size(); index++) {
            JsonNode item = hyperlinks.get(index);
            if (item.path("row").asInt(-1) == coordinate.row() && item.path("column").asInt(-1) == coordinate.column()) {
                hyperlinks.remove(index);
                return true;
            }
        }
        return false;
    }

    private int findSheetIndex(ObjectNode root, String id) {
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        for (int index = 0; index < sheets.size(); index++) if (id.equals(sheets.get(index).path("id").asText())) return index;
        return -1;
    }

    private int dimension(JsonNode value, String label) {
        if (value == null || !value.canConvertToInt() || value.intValue() < 1
                || value.intValue() > ("rowCount".equals(label) ? SnapshotMutationSupport.MAX_ROW + 1 : SnapshotMutationSupport.MAX_COLUMN + 1)) {
            throw ServiceException.validation(label + " is invalid");
        }
        return value.intValue();
    }

    private void requireSheetShape(ObjectNode sheet) {
        if (sheet.path("name").asText().isBlank() || !sheet.path("cells").isObject()
                || !sheet.path("merges").isArray() || !sheet.path("pivots").isArray() || !sheet.path("sparklines").isArray()
                || !sheet.path("drawings").isArray() || !sheet.path("drawingPayloads").isObject()) {
            throw ServiceException.validation("Restored sheet is not canonical");
        }
        dimension(sheet.get("rowCount"), "rowCount");
        dimension(sheet.get("columnCount"), "columnCount");
    }

    private void removeSheetScopedDocuments(ObjectNode root, String sheetId) {
        JsonNode names = root.get("definedNameModels");
        if (names != null && names.isArray()) removeMatching((ArrayNode) names,
                item -> "sheet".equals(item.path("scope").asText()) && sheetId.equals(item.path("sheetId").asText()));
        JsonNode documents = root.get("printDocuments");
        if (documents != null && documents.isArray()) removeMatching((ArrayNode) documents, item -> sheetId.equals(item.path("sheetId").asText()));
    }

    private void removeMatching(ArrayNode values, java.util.function.Predicate<JsonNode> predicate) {
        for (int index = values.size() - 1; index >= 0; index--) {
            if (predicate.test(values.get(index))) values.remove(index);
        }
    }

    private void replaceSheetReferences(JsonNode value, String sourceSheetId, String targetSheetId) {
        if (value.isObject()) {
            ObjectNode object = (ObjectNode) value;
            object.fields().forEachRemaining(entry -> {
                if ("sheetId".equals(entry.getKey()) && entry.getValue().isTextual() && sourceSheetId.equals(entry.getValue().asText())) {
                    object.put(entry.getKey(), targetSheetId);
                } else {
                    replaceSheetReferences(entry.getValue(), sourceSheetId, targetSheetId);
                }
            });
        } else if (value.isArray()) {
            for (JsonNode item : value) replaceSheetReferences(item, sourceSheetId, targetSheetId);
        }
    }

    private void rewriteSheetReferences(ObjectNode root, String previousName, String nextName) {
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            ObjectNode sheet = (ObjectNode) rawSheet;
            ObjectNode cells = SnapshotMutationSupport.cells(sheet);
            cells.fields().forEachRemaining(row -> {
                if (!row.getValue().isObject()) return;
                ((ObjectNode) row.getValue()).fields().forEachRemaining(cellEntry -> {
                    if (!cellEntry.getValue().isObject()) return;
                    ObjectNode cell = (ObjectNode) cellEntry.getValue();
                    JsonNode formula = cell.get("formula");
                    if (formula != null && formula.isTextual()) {
                        cell.put("formula", FormulaReferenceTransformer.renameSheet(formula.asText(), previousName, nextName));
                    }
                });
            });
        }
        JsonNode names = root.get("definedNameModels");
        if (names != null && names.isArray()) {
            for (JsonNode raw : names) {
                if (raw.isObject() && raw.path("formula").isTextual()) {
                    ((ObjectNode) raw).put("formula", FormulaReferenceTransformer.renameSheet(raw.path("formula").asText(), previousName, nextName));
                }
            }
        }
        JsonNode legacy = root.get("definedNames");
        if (legacy != null && legacy.isObject()) {
            ((ObjectNode) legacy).fields().forEachRemaining(entry -> {
                if (entry.getValue().isTextual()) ((ObjectNode) legacy).put(entry.getKey(), FormulaReferenceTransformer.renameSheet(entry.getValue().asText(), previousName, nextName));
            });
        }
    }
}
