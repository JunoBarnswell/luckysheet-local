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
            "table.add", "table.remove", "name.set", "name.remove", "workbook.calculation.mode.set",
            "pageLayout.margins.set", "pageLayout.orientation.set", "pageLayout.paperSize.set", "pageLayout.pageSetupDetail.set",
            "pageLayout.scaleToFit.set", "pageLayout.printTitles.set", "pageLayout.printArea.set", "pageLayout.printArea.clear",
            "pageLayout.pageBreak.insert", "pageLayout.pageBreak.remove", "pageLayout.pageBreak.clear",
            "pageLayout.printGridlines.set", "pageLayout.printHeadings.set", "pageLayout.viewGridlines.set", "pageLayout.viewHeadings.set"
    );

    WorkbookStateMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR);
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported workbook state mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        if (id().startsWith("pageLayout.")) {
            validatePageLayoutMutation(root, mutation);
            ObjectNode params = SnapshotMutationSupport.params(mutation);
            if (id().equals("pageLayout.printArea.set")) return List.of(SnapshotMutationSupport.range(root, params.get("range")));
            return List.of();
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
            case "workbook.calculation.mode.set" -> setCalculationMode(root, SnapshotMutationSupport.params(mutation));
            case "pageLayout.margins.set", "pageLayout.orientation.set", "pageLayout.paperSize.set", "pageLayout.pageSetupDetail.set",
                    "pageLayout.scaleToFit.set", "pageLayout.printTitles.set", "pageLayout.printArea.set", "pageLayout.printArea.clear",
                    "pageLayout.pageBreak.insert", "pageLayout.pageBreak.remove", "pageLayout.pageBreak.clear",
                    "pageLayout.printGridlines.set", "pageLayout.printHeadings.set", "pageLayout.viewGridlines.set", "pageLayout.viewHeadings.set" -> applyPageLayout(root, mutation);
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

    private void setCalculationMode(ObjectNode root, ObjectNode params) {
        String mode = SnapshotMutationSupport.text(params, "mode");
        if (!Set.of("automatic", "manual", "partial").contains(mode)) throw ServiceException.validation("Workbook calculation mode is invalid");
        ObjectNode settings = root.with("calculationSettings");
        if (!settings.has("iterativeCalculation")) settings.put("iterativeCalculation", false);
        if (!settings.has("maximumIterations")) settings.put("maximumIterations", 100);
        if (!settings.has("maximumChange")) settings.put("maximumChange", 0.001);
        if (!settings.has("precisionAsDisplayed")) settings.put("precisionAsDisplayed", false);
        if (!settings.has("calculateBeforeSave")) settings.put("calculateBeforeSave", true);
        if (!settings.has("fullCalculationOnLoad")) settings.put("fullCalculationOnLoad", false);
        settings.put("mode", mode);
    }

    private void validatePageLayoutMutation(ObjectNode root, OperationMutation mutation) {
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        String sheetId = SnapshotMutationSupport.text(params, "sheetId");
        if (!mutation.sheetId().equals(sheetId)) throw ServiceException.validation("Page layout mutation sheet does not match operation sheet");
        SnapshotMutationSupport.sheet(root, sheetId);
        switch (id()) {
            case "pageLayout.margins.set" -> {
                ObjectNode margins = SnapshotMutationSupport.requiredObject(params, "margins");
                for (String field : List.of("top", "right", "bottom", "left", "header", "footer")) finiteNonNegative(margins.get(field), "Page layout margin " + field);
            }
            case "pageLayout.orientation.set" -> {
                String orientation = SnapshotMutationSupport.text(params, "orientation");
                if (!Set.of("portrait", "landscape").contains(orientation)) throw ServiceException.validation("Page layout orientation is invalid");
            }
            case "pageLayout.paperSize.set" -> {
                String paper = SnapshotMutationSupport.text(params, "paperSize");
                if (!Set.of("letter", "a4", "a3", "legal", "custom").contains(paper)) throw ServiceException.validation("Page layout paper size is invalid");
            }
            case "pageLayout.pageSetupDetail.set" -> validatePageSetup(SnapshotMutationSupport.requiredObject(params, "pageSetup"));
            case "pageLayout.scaleToFit.set" -> {
                JsonNode scale = params.get("scale");
                if (scale == null || !scale.isNumber() || scale.asDouble() <= 0 || scale.asDouble() > 400) throw ServiceException.validation("Page layout scale is invalid");
                optionalPositiveInteger(params.get("fitToWidth"), "Page layout fitToWidth");
                optionalPositiveInteger(params.get("fitToHeight"), "Page layout fitToHeight");
            }
            case "pageLayout.printTitles.set" -> {
                validateSpan(params.get("repeatRows"), "Page layout repeatRows");
                validateSpan(params.get("repeatColumns"), "Page layout repeatColumns");
            }
            case "pageLayout.printArea.set" -> {
                RangeRef range = SnapshotMutationSupport.range(root, params.get("range"));
                SnapshotMutationSupport.requireSheet(range, sheetId);
            }
            case "pageLayout.pageBreak.insert", "pageLayout.pageBreak.remove" -> validatePageBreak(params.get("pageBreak"), sheetId);
            case "pageLayout.printGridlines.set", "pageLayout.printHeadings.set", "pageLayout.viewGridlines.set", "pageLayout.viewHeadings.set" -> {
                if (!params.path("enabled").isBoolean()) throw ServiceException.validation("Page layout toggle enabled is required");
            }
            case "pageLayout.printArea.clear" -> validatePrintAreas(root, params.get("printAreas"), sheetId);
            case "pageLayout.pageBreak.clear" -> validatePageBreaks(params.get("pageBreaks"), sheetId);
            default -> throw ServiceException.validation("Unsupported page layout mutation: " + id());
        }
    }

    private void validatePageSetup(ObjectNode setup) {
        String paper = SnapshotMutationSupport.text(setup, "paperSize");
        if (!Set.of("letter", "a4", "a3", "legal", "custom").contains(paper)) throw ServiceException.validation("Print paper size is invalid");
        String orientation = SnapshotMutationSupport.text(setup, "orientation");
        if (!Set.of("portrait", "landscape").contains(orientation)) throw ServiceException.validation("Print orientation is invalid");
        ObjectNode margins = SnapshotMutationSupport.requiredObject(setup, "margins");
        for (String field : List.of("top", "right", "bottom", "left", "header", "footer")) finiteNonNegative(margins.get(field), "Print margin " + field);
        JsonNode scale = setup.get("scale");
        if (scale == null || !scale.isNumber() || scale.asDouble() <= 0 || scale.asDouble() > 400) throw ServiceException.validation("Print scale is invalid");
        for (String field : List.of("printGridlines", "printHeadings", "centerHorizontally", "centerVertically")) if (!setup.path(field).isBoolean()) throw ServiceException.validation("Print page setup " + field + " must be boolean");
    }

    private void validateSpan(JsonNode value, String label) {
        if (value == null || value.isNull()) return;
        if (!value.isObject() || !value.path("start").canConvertToInt() || !value.path("end").canConvertToInt() || value.path("start").asInt() < 0 || value.path("end").asInt() < value.path("start").asInt()) throw ServiceException.validation(label + " is invalid");
    }

    private void optionalPositiveInteger(JsonNode value, String label) {
        if (value == null || value.isNull()) return;
        if (!value.isIntegralNumber() || value.asInt() < 1) throw ServiceException.validation(label + " is invalid");
    }

    private void validatePrintAreas(ObjectNode root, JsonNode value, String sheetId) {
        if (value == null) return;
        if (!value.isArray()) throw ServiceException.validation("Page layout printAreas must be an array");
        for (JsonNode area : value) {
            if (!area.isObject() || !sheetId.equals(SnapshotMutationSupport.text((ObjectNode) area, "sheetId"))) throw ServiceException.validation("Page layout print area sheet is invalid");
            RangeRef range = SnapshotMutationSupport.range(root, area.get("range"));
            SnapshotMutationSupport.requireSheet(range, sheetId);
        }
    }

    private void validatePageBreaks(JsonNode value, String sheetId) {
        if (value == null) return;
        if (!value.isArray()) throw ServiceException.validation("Page layout pageBreaks must be an array");
        for (JsonNode pageBreak : value) validatePageBreak(pageBreak, sheetId);
    }

    private ObjectNode printDocumentFor(ObjectNode root, String sheetId) {
        ArrayNode documents = SnapshotMutationSupport.array(root, "printDocuments");
        for (JsonNode document : documents) if (sheetId.equals(document.path("sheetId").asText()) && document.isObject()) return (ObjectNode) document;
        ObjectNode document = JsonNodeFactory.instance.objectNode();
        document.put("schema", "PrintDocument");
        document.put("unitId", SnapshotMutationSupport.text(root, "unitId"));
        document.put("sheetId", sheetId);
        ObjectNode setup = document.putObject("pageSetup");
        setup.put("paperSize", "a4");
        setup.put("orientation", "portrait");
        ObjectNode margins = setup.putObject("margins");
        margins.put("top", 72).put("right", 72).put("bottom", 72).put("left", 72).put("header", 36).put("footer", 36);
        setup.put("scale", 100).put("printGridlines", false).put("printHeadings", false).put("centerHorizontally", false).put("centerVertically", false);
        document.putArray("printAreas");
        document.putArray("pageBreaks");
        documents.add(document);
        return document;
    }

    private void applyPageLayout(ObjectNode root, OperationMutation mutation) {
        validatePageLayoutMutation(root, mutation);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        String sheetId = SnapshotMutationSupport.text(params, "sheetId");
        ObjectNode document = printDocumentFor(root, sheetId);
        ObjectNode setup = SnapshotMutationSupport.requiredObject(document, "pageSetup");
        switch (id()) {
            case "pageLayout.margins.set" -> setup.set("margins", params.get("margins").deepCopy());
            case "pageLayout.orientation.set" -> setup.set("orientation", params.get("orientation"));
            case "pageLayout.paperSize.set" -> setup.set("paperSize", params.get("paperSize"));
            case "pageLayout.pageSetupDetail.set" -> document.set("pageSetup", params.get("pageSetup").deepCopy());
            case "pageLayout.scaleToFit.set" -> { setup.set("scale", params.get("scale")); copyOrRemove(setup, params, "fitToWidth"); copyOrRemove(setup, params, "fitToHeight"); }
            case "pageLayout.printTitles.set" -> { if (params.has("repeatRows")) copyOrRemove(document, params, "repeatRows"); if (params.has("repeatColumns")) copyOrRemove(document, params, "repeatColumns"); }
            case "pageLayout.printArea.set" -> { ArrayNode areas = document.putArray("printAreas"); ObjectNode area = areas.addObject(); area.put("sheetId", sheetId); area.set("range", params.get("range").deepCopy()); }
            case "pageLayout.printArea.clear" -> {
                if (params.has("printAreas")) document.set("printAreas", params.get("printAreas").deepCopy());
                else document.putArray("printAreas");
            }
            case "pageLayout.pageBreak.insert" -> { ArrayNode breaks = document.withArray("pageBreaks"); removeBreak(breaks, params.get("pageBreak")); breaks.add(params.get("pageBreak").deepCopy()); }
            case "pageLayout.pageBreak.remove" -> removeBreak(document.withArray("pageBreaks"), params.get("pageBreak"));
            case "pageLayout.pageBreak.clear" -> {
                if (params.has("pageBreaks")) document.set("pageBreaks", params.get("pageBreaks").deepCopy());
                else document.putArray("pageBreaks");
            }
            case "pageLayout.printGridlines.set" -> setup.set("printGridlines", params.get("enabled"));
            case "pageLayout.printHeadings.set" -> setup.set("printHeadings", params.get("enabled"));
            case "pageLayout.viewGridlines.set" -> SnapshotMutationSupport.sheet(root, sheetId).set("showGridlines", params.get("enabled"));
            case "pageLayout.viewHeadings.set" -> SnapshotMutationSupport.sheet(root, sheetId).set("showHeaders", params.get("enabled"));
            default -> throw ServiceException.validation("Unsupported page layout mutation: " + id());
        }
    }

    private void copyOrRemove(ObjectNode target, ObjectNode source, String field) {
        JsonNode value = source.get(field);
        if (value == null || value.isNull()) target.remove(field); else target.set(field, value.deepCopy());
    }

    private void removeBreak(ArrayNode breaks, JsonNode candidate) {
        for (int index = breaks.size() - 1; index >= 0; index--) {
            JsonNode current = breaks.get(index);
            if (java.util.Objects.equals(current.get("row"), candidate.get("row")) && java.util.Objects.equals(current.get("column"), candidate.get("column"))) breaks.remove(index);
        }
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

}
