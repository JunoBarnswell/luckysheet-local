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

/** Reducers for workbook-owned tables, names, and persisted print documents. */
final class WorkbookStateMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "table.add", "table.remove", "name.set", "name.remove",
            "print.pageSetup.set", "print.area.set", "print.area.clear", "print.pageBreak.set",
            "print.pageBreak.remove", "print.pageBreaks.clear", "print.document.replace"
    );

    WorkbookStateMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, checksProtection(id), action(id));
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported workbook state mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        if (id().startsWith("print.")) {
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            ObjectNode document = SnapshotMutationSupport.requiredObject(params, "document");
            validateDocument(root, mutation.sheetId(), document);
            return List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
        }
        if (id().equals("table.add")) {
            ObjectNode table = requiredMutationObject(mutation);
            JsonNode sourceRange = table.get("sourceRange");
            return sourceRange == null || sourceRange.isNull() ? List.of() : List.of(SnapshotMutationSupport.range(root, sourceRange));
        }
        if (id().equals("table.remove")) {
            String tableId = rawId(mutation.params(), "Table id");
            ObjectNode table = SnapshotMutationSupport.findById(SnapshotMutationSupport.dataModelArray(root, "tables"), tableId);
            return table == null || table.get("sourceRange") == null ? List.of() : List.of(SnapshotMutationSupport.range(root, table.get("sourceRange")));
        }
        SnapshotMutationSupport.params(mutation);
        return List.of();
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        switch (id()) {
            case "table.add" -> addTable(root, requiredMutationObject(mutation));
            case "table.remove" -> removeTable(root, rawId(mutation.params(), "Table id"));
            case "name.set" -> setName(root, SnapshotMutationSupport.params(mutation));
            case "name.remove" -> removeName(root, SnapshotMutationSupport.params(mutation));
            case "print.pageSetup.set", "print.area.set", "print.area.clear", "print.pageBreak.set", "print.pageBreak.remove", "print.pageBreaks.clear", "print.document.replace" -> setDocument(root, mutation.sheetId(), SnapshotMutationSupport.params(mutation));
            default -> throw ServiceException.validation("Unsupported workbook state mutation: " + id());
        }
        return root;
    }

    private void addTable(ObjectNode root, ObjectNode table) {
        validateTable(root, table);
        ArrayNode tables = SnapshotMutationSupport.dataModelArray(root, "tables");
        String id = SnapshotMutationSupport.text(table, "id");
        if (SnapshotMutationSupport.findById(tables, id) != null) throw ServiceException.conflict("Workbook table already exists: " + id);
        tables.add(table.deepCopy());
    }

    private void removeTable(ObjectNode root, String tableId) {
        if (!SnapshotMutationSupport.removeById(SnapshotMutationSupport.dataModelArray(root, "tables"), tableId)) throw ServiceException.notFound("Workbook table not found: " + tableId);
    }

    private void validateTable(ObjectNode root, ObjectNode table) {
        SnapshotMutationSupport.text(table, "id");
        String name = SnapshotMutationSupport.text(table, "name");
        if (name.length() > 255) throw ServiceException.validation("Workbook table name is too long");
        JsonNode rowCount = table.get("rowCount");
        JsonNode blockSize = table.get("blockSize");
        JsonNode revision = table.get("revision");
        if (rowCount == null || !rowCount.isIntegralNumber() || rowCount.longValue() < 0) throw ServiceException.validation("Workbook table rowCount is invalid");
        if (blockSize == null || !blockSize.isIntegralNumber() || blockSize.intValue() < 1) throw ServiceException.validation("Workbook table blockSize is invalid");
        if (revision == null || !revision.isIntegralNumber() || revision.longValue() < 0) throw ServiceException.validation("Workbook table revision is invalid");
        if (!table.path("fields").isArray() || !table.path("blocks").isArray()) throw ServiceException.validation("Workbook table fields and blocks are required");
        String sourceSheetId = SnapshotMutationSupport.optionalText(table, "sourceSheetId");
        JsonNode sourceRange = table.get("sourceRange");
        if (sourceRange != null && !sourceRange.isNull()) {
            RangeRef range = SnapshotMutationSupport.range(root, sourceRange);
            if (sourceSheetId != null && !sourceSheetId.equals(range.sheetId())) throw ServiceException.validation("Workbook table source sheet does not match source range");
        } else if (sourceSheetId != null) {
            SnapshotMutationSupport.sheet(root, sourceSheetId);
        }
    }

    private void setName(ObjectNode root, ObjectNode params) {
        ObjectNode model = SnapshotMutationSupport.requiredObject(params, "model");
        String name = SnapshotMutationSupport.text(model, "name");
        String formula = SnapshotMutationSupport.text(model, "formula");
        String scope = SnapshotMutationSupport.text(model, "scope");
        if (!scope.equals("workbook") && !scope.equals("sheet")) throw ServiceException.validation("Defined name scope is invalid");
        String sheetId = SnapshotMutationSupport.optionalText(model, "sheetId");
        if (scope.equals("sheet") && (sheetId == null || sheetId.isBlank())) throw ServiceException.validation("Sheet-scoped name requires sheetId");
        if (sheetId != null) SnapshotMutationSupport.sheet(root, sheetId);
        if (name.length() > 255 || formula.length() > 32_767) throw ServiceException.validation("Defined name is too large");
        ArrayNode models = SnapshotMutationSupport.array(root, "definedNameModels");
        int existing = nameIndex(models, name, scope, sheetId);
        if (existing >= 0) models.set(existing, model.deepCopy());
        else models.add(model.deepCopy());
        ObjectNode legacyFormulaView = SnapshotMutationSupport.object(root, "definedNames");
        if (scope.equals("workbook")) legacyFormulaView.put(name, formula);
    }

    private void removeName(ObjectNode root, ObjectNode params) {
        String name = SnapshotMutationSupport.text(params, "name");
        String scope = SnapshotMutationSupport.optionalText(params, "scope");
        if (scope == null) scope = "workbook";
        if (!scope.equals("workbook") && !scope.equals("sheet")) throw ServiceException.validation("Defined name scope is invalid");
        String sheetId = SnapshotMutationSupport.optionalText(params, "sheetId");
        ArrayNode models = SnapshotMutationSupport.array(root, "definedNameModels");
        int index = nameIndex(models, name, scope, sheetId);
        if (index < 0) throw ServiceException.notFound("Defined name not found: " + name);
        models.remove(index);
        if (scope.equals("workbook")) SnapshotMutationSupport.object(root, "definedNames").remove(name);
    }

    private int nameIndex(ArrayNode models, String name, String scope, String sheetId) {
        for (int index = 0; index < models.size(); index++) {
            JsonNode model = models.get(index);
            if (model.isObject()
                    && name.equalsIgnoreCase(model.path("name").asText())
                    && scope.equals(model.path("scope").asText())
                    && java.util.Objects.equals(sheetId, model.path("sheetId").isMissingNode() ? null : model.path("sheetId").asText())) {
                return index;
            }
        }
        return -1;
    }

    private void setDocument(ObjectNode root, String sheetId, ObjectNode params) {
        ObjectNode document = SnapshotMutationSupport.requiredObject(params, "document");
        validateDocument(root, sheetId, document);
        ArrayNode documents = SnapshotMutationSupport.array(root, "printDocuments");
        for (int index = 0; index < documents.size(); index++) {
            if (sheetId.equals(documents.get(index).path("sheetId").asText())) {
                documents.set(index, document.deepCopy());
                return;
            }
        }
        documents.add(document.deepCopy());
    }

    private void validateDocument(ObjectNode root, String expectedSheetId, ObjectNode document) {
        if (!"PrintDocument".equals(SnapshotMutationSupport.text(document, "schema"))) throw ServiceException.validation("Print document schema is invalid");
        if (!SnapshotMutationSupport.text(root, "unitId").equals(SnapshotMutationSupport.text(document, "unitId"))) throw ServiceException.validation("Print document unit does not match workbook");
        SnapshotMutationSupport.requireEntitySheet(document, expectedSheetId);
        ObjectNode setup = SnapshotMutationSupport.requiredObject(document, "pageSetup");
        String paper = SnapshotMutationSupport.text(setup, "paperSize");
        if (!Set.of("letter", "a4", "a3", "legal", "custom").contains(paper)) throw ServiceException.validation("Print paper size is invalid");
        String orientation = SnapshotMutationSupport.text(setup, "orientation");
        if (!orientation.equals("portrait") && !orientation.equals("landscape")) throw ServiceException.validation("Print orientation is invalid");
        ObjectNode margins = SnapshotMutationSupport.requiredObject(setup, "margins");
        for (String field : List.of("top", "right", "bottom", "left", "header", "footer")) finiteNonNegative(margins.get(field), "Print margin " + field);
        JsonNode scale = setup.get("scale");
        if (scale == null || !scale.isNumber() || scale.asDouble() <= 0 || scale.asDouble() > 400) throw ServiceException.validation("Print scale is invalid");
        for (String field : List.of("printGridlines", "printHeadings", "centerHorizontally", "centerVertically")) {
            if (!setup.path(field).isBoolean()) throw ServiceException.validation("Print page setup " + field + " must be boolean");
        }
        ArrayNode areas = SnapshotMutationSupport.requiredArray(document, "printAreas");
        for (JsonNode area : areas) {
            if (!area.isObject() || !expectedSheetId.equals(area.path("sheetId").asText())) throw ServiceException.validation("Print area targets another sheet");
            RangeRef range = SnapshotMutationSupport.range(root, area.get("range"));
            SnapshotMutationSupport.requireSheet(range, expectedSheetId);
        }
        ArrayNode breaks = SnapshotMutationSupport.requiredArray(document, "pageBreaks");
        for (JsonNode pageBreak : breaks) validatePageBreak(pageBreak, expectedSheetId);
    }

    private void validatePageBreak(JsonNode pageBreak, String sheetId) {
        if (pageBreak == null || !pageBreak.isObject() || !sheetId.equals(pageBreak.path("sheetId").asText())) throw ServiceException.validation("Print page break targets another sheet");
        JsonNode row = pageBreak.get("row");
        JsonNode column = pageBreak.get("column");
        boolean hasRow = row != null && !row.isNull();
        boolean hasColumn = column != null && !column.isNull();
        if (hasRow == hasColumn) throw ServiceException.validation("Print page break must specify exactly one axis");
        JsonNode value = hasRow ? row : column;
        if (!value.isIntegralNumber() || value.intValue() < 0) throw ServiceException.validation("Print page break is invalid");
    }

    private ObjectNode requiredMutationObject(OperationMutation mutation) {
        if (mutation.params() == null || !mutation.params().isObject()) throw ServiceException.validation("Mutation params must be an object");
        return (ObjectNode) mutation.params();
    }

    private String rawId(JsonNode value, String label) {
        if (value == null || !value.isTextual() || value.asText().isBlank()) throw ServiceException.validation(label + " is required");
        return value.asText();
    }

    private void finiteNonNegative(JsonNode value, String label) {
        if (value == null || !value.isNumber() || !Double.isFinite(value.asDouble()) || value.asDouble() < 0) throw ServiceException.validation(label + " is invalid");
    }

    private static boolean checksProtection(String id) {
        return id.equals("table.add") || id.equals("table.remove");
    }

    private static String action(String id) {
        return id.startsWith("print.") ? "print" : id.startsWith("table.") ? "structure" : "edit-cell";
    }
}
