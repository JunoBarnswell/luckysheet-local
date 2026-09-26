package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookSnapshotValidator;
import com.xc.luckysheet.server.service.ServiceException;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/** Canonical snapshot reducers for worksheet lifecycle and persisted hyperlinks. */
final class WorkbookStructureMutationDescriptor extends CanonicalJsonMutationDescriptor {
    private static final List<String> CHART_TEXT_FORMULA_FIELDS = List.of(
            "titleText.linkedFormula", "legend.text.linkedFormula",
            "categoryAxis.titleText.linkedFormula", "valueAxis.titleText.linkedFormula",
            "secondaryCategoryAxis.titleText.linkedFormula", "secondaryValueAxis.titleText.linkedFormula",
            "dataTable.font.linkedFormula");
    static final Set<String> IDS = Set.of(
            "sheet.add", "sheet.remove", "sheet.rename", "sheet.duplicated", "sheet.restore",
            "hyperlink.set", "hyperlink.remove"
    );

    WorkbookStructureMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR);
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
        return applyWithPatch(snapshot, mutation).snapshot();
    }

    @Override
    public MutationApplication applyWithPatch(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        StructuralPatch structuralPatch = null;
        switch (id()) {
            case "sheet.add" -> add(root, mutation.sheetId(), params);
            case "sheet.remove" -> remove(root, params);
            case "sheet.rename" -> structuralPatch = rename(root, mutation.sheetId(), params);
            case "sheet.duplicated" -> duplicate(root, params);
            case "sheet.restore" -> restore(root, params);
            case "hyperlink.set" -> setHyperlink(root, mutation.sheetId(), params);
            case "hyperlink.remove" -> removeHyperlink(root, mutation.sheetId(), params);
            default -> throw ServiceException.validation("Unsupported workbook structure mutation: " + id());
        }
        return new MutationApplication(root, structuralPatch);
    }

    private void add(ObjectNode root, String mutationSheetId, ObjectNode params) {
        String id = SnapshotMutationSupport.text(params, "id");
        String name = SnapshotMutationSupport.text(params, "name").trim();
        if (!mutationSheetId.equals(id) || name.isBlank()) throw ServiceException.validation("sheet.add identity is invalid");
        if (findSheetIndex(root, id) >= 0) throw ServiceException.conflict("Sheet already exists: " + id);
        int rows = params.has("rowCount") ? dimension(params.get("rowCount"), "rowCount") : 1000;
        int columns = params.has("columnCount") ? dimension(params.get("columnCount"), "columnCount") : 26;
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        WorkbookSnapshotValidator.requireWorksheetNameAvailable(sheets, name, null);
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
        ObjectNode review = sheet.putObject("review");
        review.putObject("notesByCell");
        review.putObject("notesById");
        review.putObject("threadIdsByCell");
        review.putObject("threadsById");
        sheets.add(sheet);
    }

    private void remove(ObjectNode root, ObjectNode params) {
        String id = SnapshotMutationSupport.text(params, "id");
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        if (sheets.size() <= 1) throw ServiceException.validation("A workbook must keep at least one worksheet");
        int index = findSheetIndex(root, id);
        if (index < 0) throw ServiceException.notFound("Sheet not found: " + id);
        validateNoExternalSheetReferences(root, id, sheets.get(index).path("name").asText());
        sheets.remove(index);
        removeSheetScopedDocuments(root, id);
    }

    private StructuralPatch rename(ObjectNode root, String mutationSheetId, ObjectNode params) {
        String sheetId = SnapshotMutationSupport.text(params, "sheetId");
        String name = SnapshotMutationSupport.text(params, "name").trim();
        if (!mutationSheetId.equals(sheetId) || name.isBlank()) throw ServiceException.validation("sheet.rename identity is invalid");
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        String previousName = sheet.path("name").asText();
        WorkbookSnapshotValidator.requireWorksheetNameAvailable(SnapshotMutationSupport.sheets(root), name, sheetId);
        StructuralPatch patch = previousName.equals(name)
                ? new StructuralPatch(StructuralPatch.VERSION, "sheet.rename", List.of(), List.of(), List.of())
                : StructuralSnapshotReducer.renameSheetReferences(root, sheetId, previousName, name);
        sheet.put("name", name);
        return patch;
    }

    private void duplicate(ObjectNode root, ObjectNode params) {
        String sourceSheetId = SnapshotMutationSupport.text(params, "sourceSheetId");
        String newId = SnapshotMutationSupport.text(params, "newId");
        String newName = SnapshotMutationSupport.text(params, "newName").trim();
        if (newName.isBlank() || findSheetIndex(root, newId) >= 0) throw ServiceException.conflict("Duplicate sheet identity is invalid");
        int sourceIndex = findSheetIndex(root, sourceSheetId);
        if (sourceIndex < 0) throw ServiceException.notFound("Sheet not found: " + sourceSheetId);
        WorkbookSnapshotValidator.requireWorksheetNameAvailable(SnapshotMutationSupport.sheets(root), newName, null);
        ObjectNode source = (ObjectNode) SnapshotMutationSupport.sheets(root).get(sourceIndex);
        String sourceName = source.path("name").asText();
        ObjectNode copy = source.deepCopy();
        remapDuplicatedSheetReferences(root, copy, sourceSheetId, sourceName, newId, newName);
        copy.put("id", newId);
        copy.put("name", newName);
        SnapshotMutationSupport.sheets(root).insert(sourceIndex + 1, copy);
        cloneSheetScopedDocuments(root, sourceSheetId, sourceName, newId, newName);
    }

    private void restore(ObjectNode root, ObjectNode params) {
        JsonNode raw = params.get("sheet");
        if (raw == null || !raw.isObject()) throw ServiceException.validation("sheet.restore requires a canonical sheet snapshot");
        ObjectNode sheet = ((ObjectNode) raw).deepCopy();
        String id = sheet.path("id").asText().trim();
        if (id.isBlank() || findSheetIndex(root, id) >= 0) throw ServiceException.conflict("Restored sheet identity is invalid");
        requireSheetShape(sheet);
        WorkbookSnapshotValidator.requireWorksheetNameAvailable(SnapshotMutationSupport.sheets(root), sheet.path("name").asText(), null);
        int index = params.path("index").isInt() ? params.path("index").intValue() : SnapshotMutationSupport.sheets(root).size();
        int bounded = Math.max(0, Math.min(index, SnapshotMutationSupport.sheets(root).size()));
        SnapshotMutationSupport.sheets(root).insert(bounded, sheet);
        restoreSheetScopedDocuments(root, sheet);
        sheet.remove("lifecycleDefinedNames");
        sheet.remove("lifecyclePrintDocument");
    }

    private void setHyperlink(ObjectNode root, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, sheetId, params);
        JsonNode hyperlink = params.get("hyperlink");
        if (hyperlink == null || !hyperlink.isObject() || hyperlink.path("id").asText().isBlank()
                || !hyperlink.path("target").isObject()) throw ServiceException.validation("hyperlink.set requires a canonical hyperlink");
        validateHyperlinkTarget(root, sheetId, hyperlink.get("target"));
        if (hyperlink.has("tooltip") && !hyperlink.get("tooltip").isTextual()) throw ServiceException.validation("Hyperlink tooltip is invalid");
        ArrayNode hyperlinks = SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, sheetId), "hyperlinks");
        removeHyperlinkAt(hyperlinks, coordinate);
        ObjectNode entry = hyperlinks.objectNode();
        entry.put("row", coordinate.row());
        entry.put("column", coordinate.column());
        entry.set("hyperlink", hyperlink.deepCopy());
        hyperlinks.add(entry);
    }

    private void validateHyperlinkTarget(ObjectNode root, String sourceSheetId, JsonNode target) {
        String kind = target.path("kind").asText().trim();
        if (kind.isBlank()) throw ServiceException.validation("Hyperlink target kind is required");
        switch (kind) {
            case "url" -> {
                String value = target.path("url").asText().trim();
                if (value.isBlank()) throw ServiceException.validation("Hyperlink URL is required");
                try {
                    URI uri = new URI(value);
                    if (!("http".equalsIgnoreCase(uri.getScheme()) || "https".equalsIgnoreCase(uri.getScheme()) || "ftp".equalsIgnoreCase(uri.getScheme()))
                            || uri.getHost() == null || uri.getHost().isBlank()) {
                        throw ServiceException.validation("Unsupported hyperlink URL scheme");
                    }
                } catch (URISyntaxException exception) {
                    throw ServiceException.validation("Invalid hyperlink URL");
                }
            }
            case "email" -> {
                String address = target.path("address").asText().trim();
                if (!address.matches("^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$")) throw ServiceException.validation("Invalid email hyperlink address");
                if (target.has("subject") && !target.get("subject").isTextual()) throw ServiceException.validation("Email hyperlink subject is invalid");
            }
            case "sheet" -> {
                String targetSheetId = target.path("sheetId").asText().trim();
                ObjectNode targetSheet = findSheet(root, targetSheetId);
                if (targetSheet == null) throw ServiceException.notFound("Hyperlink target sheet not found: " + targetSheetId);
                boolean hasAddress = target.has("address");
                boolean hasRow = target.has("row");
                boolean hasColumn = target.has("column");
                if (hasAddress && (hasRow || hasColumn) || !hasAddress && !(hasRow && hasColumn)) {
                    throw ServiceException.validation("Worksheet hyperlink address must be canonical");
                }
                int row;
                int column;
                if (hasAddress) {
                    String address = target.path("address").asText().trim();
                    java.util.regex.Matcher matcher = java.util.regex.Pattern.compile("^([A-Za-z]+)([1-9][0-9]*)$").matcher(address);
                    if (!matcher.matches()) throw ServiceException.validation("Worksheet hyperlink address is invalid");
                    column = columnIndex(matcher.group(1));
                    try {
                        row = Math.subtractExact(Integer.parseInt(matcher.group(2)), 1);
                    } catch (NumberFormatException | ArithmeticException exception) {
                        throw ServiceException.validation("Worksheet hyperlink row is invalid");
                    }
                } else {
                    if (!target.path("row").canConvertToInt() || !target.path("column").canConvertToInt()) throw ServiceException.validation("Worksheet hyperlink coordinates are invalid");
                    row = target.path("row").intValue();
                    column = target.path("column").intValue();
                }
                if (row < 0 || column < 0 || row >= targetSheet.path("rowCount").asInt(-1) || column >= targetSheet.path("columnCount").asInt(-1)) {
                    throw ServiceException.validation("Worksheet hyperlink address is outside the worksheet bounds");
                }
            }
            case "name" -> {
                String name = target.path("name").asText().trim();
                if (!name.matches("^[A-Za-z_\\\\][A-Za-z0-9_.]*$")) throw ServiceException.validation("Invalid defined-name hyperlink");
                JsonNode names = root.get("definedNameModels");
                boolean found = names != null && names.isArray();
                if (found) {
                    found = false;
                    for (JsonNode entry : names) {
                        if (name.equalsIgnoreCase(entry.path("name").asText())
                                && ("workbook".equals(entry.path("scope").asText()) || sourceSheetId.equals(entry.path("sheetId").asText()))) {
                            found = true;
                            break;
                        }
                    }
                }
                if (!found) throw ServiceException.notFound("Defined name not found: " + name);
            }
            default -> throw ServiceException.validation("Unsupported hyperlink target kind: " + kind);
        }
    }

    private ObjectNode findSheet(ObjectNode root, String sheetId) {
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        for (JsonNode sheet : sheets) if (sheetId.equals(sheet.path("id").asText()) && sheet.isObject()) return (ObjectNode) sheet;
        return null;
    }

    private int columnIndex(String label) {
        long value = 0;
        for (int index = 0; index < label.length(); index++) value = value * 26 + Character.toUpperCase(label.charAt(index)) - 'A' + 1L;
        if (value < 1 || value > Integer.MAX_VALUE) throw ServiceException.validation("Worksheet hyperlink column is invalid");
        return (int) value - 1;
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
        if (!sheet.path("id").isTextual() || !sheet.path("id").asText().equals(sheet.path("id").asText().trim())
                || !sheet.path("name").isTextual() || sheet.path("name").asText().isBlank() || !sheet.path("cells").isObject()
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

    private void cloneSheetScopedDocuments(ObjectNode root, String sourceSheetId, String sourceName, String targetSheetId, String targetName) {
        JsonNode names = root.get("definedNameModels");
        if (names != null && names.isArray()) {
            List<JsonNode> copies = new ArrayList<>();
            for (JsonNode raw : names) {
                if (!raw.isObject() || !"sheet".equals(raw.path("scope").asText()) || !sourceSheetId.equals(raw.path("sheetId").asText())) continue;
                ObjectNode copy = ((ObjectNode) raw).deepCopy();
                copy.put("sheetId", targetSheetId);
                if (copy.path("formula").isTextual()) copy.put("formula", FormulaReferenceTransformer.renameSheet(copy.path("formula").asText(), sourceName, targetName));
                JsonNode anchor = copy.get("anchor");
                if (anchor != null && anchor.isObject() && sourceSheetId.equals(anchor.path("sheetId").asText())) ((ObjectNode) anchor).put("sheetId", targetSheetId);
                copies.add(copy);
            }
            for (JsonNode copy : copies) ((ArrayNode) names).add(copy);
        }
        JsonNode documents = root.get("printDocuments");
        if (documents != null && documents.isArray()) {
            List<JsonNode> copies = new ArrayList<>();
            for (JsonNode raw : documents) {
                if (!raw.isObject() || !sourceSheetId.equals(raw.path("sheetId").asText())) continue;
                ObjectNode copy = ((ObjectNode) raw).deepCopy();
                copy.put("sheetId", targetSheetId);
                remapPrintDocument(copy, sourceSheetId, targetSheetId);
                copies.add(copy);
            }
            for (JsonNode copy : copies) ((ArrayNode) documents).add(copy);
        }
    }

    private void restoreSheetScopedDocuments(ObjectNode root, ObjectNode sheet) {
        String sheetId = sheet.path("id").asText();
        JsonNode names = sheet.get("lifecycleDefinedNames");
        JsonNode rootNames = root.get("definedNameModels");
        if (names != null && names.isArray() && rootNames != null && rootNames.isArray()) {
            for (JsonNode raw : names) if (raw.isObject() && sheetId.equals(raw.path("sheetId").asText())) ((ArrayNode) rootNames).add(raw.deepCopy());
        }
        JsonNode printDocument = sheet.get("lifecyclePrintDocument");
        JsonNode documents = root.get("printDocuments");
        if (printDocument != null && printDocument.isObject() && documents != null && documents.isArray() && sheetId.equals(printDocument.path("sheetId").asText())) ((ArrayNode) documents).add(printDocument.deepCopy());
    }

    private void remapPrintDocument(ObjectNode document, String sourceSheetId, String targetSheetId) {
        JsonNode areas = document.get("printAreas");
        if (areas != null && areas.isArray()) for (JsonNode area : areas) if (area.isObject()) {
            ((ObjectNode) area).put("sheetId", targetSheetId);
            remapRange(area.get("range"), sourceSheetId, targetSheetId);
        }
        JsonNode breaks = document.get("pageBreaks");
        if (breaks != null && breaks.isArray()) for (JsonNode pageBreak : breaks) if (pageBreak.isObject()) ((ObjectNode) pageBreak).put("sheetId", targetSheetId);
    }

    private void validateNoExternalSheetReferences(ObjectNode root, String sourceSheetId, String sourceName) {
        List<String> references = new ArrayList<>();
        ObjectNode deletedSheet = SnapshotMutationSupport.sheet(root, sourceSheetId);
        Set<String> deletedPivotIds = collectIds(deletedSheet.get("pivots"));
        Set<String> deletedSheetTableIds = collectIds(deletedSheet.get("sheetTables"));
        Set<String> deletedDrawingIds = collectIds(deletedSheet.get("drawings"));
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            ObjectNode sheet = (ObjectNode) rawSheet;
            if (sourceSheetId.equals(sheet.path("id").asText())) continue;
            JsonNode cells = sheet.get("cells");
            if (cells != null && cells.isObject()) cells.fields().forEachRemaining(row -> {
                if (!row.getValue().isObject()) return;
                row.getValue().fields().forEachRemaining(cell -> {
                    if (!cell.getValue().isObject()) return;
                    JsonNode formula = cell.getValue().get("formula");
                    addFormulaReference(references, "cell-formula", formula, sourceName);
                    addFormulaReference(references, "cell-source-formula", cell.getValue().path("formulaMetadata").get("sourceFormula"), sourceName);
                    JsonNode presentation = cell.getValue().get("presentation");
                    if (presentation != null && "barcode".equals(presentation.path("kind").asText())) {
                        JsonNode barcodeSource = presentation.get("source");
                        if (barcodeSource != null && "formula".equals(barcodeSource.path("kind").asText())) {
                            addFormulaReference(references, "barcode-formula", barcodeSource.get("formula"), sourceName);
                        }
                    }
                });
            });
            JsonNode tableSheet = sheet.get("tableSheet");
            if (tableSheet != null && tableSheet.isObject()) {
                JsonNode columns = optionalArray(tableSheet, "columns");
                if (columns != null) for (JsonNode column : columns) {
                    addFormulaReference(references, "table-sheet-formula:" + column.path("fieldId").asText(), column.get("formula"), sourceName);
                }
            }
            JsonNode hyperlinks = sheet.get("hyperlinks");
            if (hyperlinks != null && hyperlinks.isArray()) for (JsonNode hyperlink : hyperlinks) if (sourceSheetId.equals(hyperlink.path("hyperlink").path("target").path("sheetId").asText())) references.add("hyperlink:" + hyperlink.path("hyperlink").path("id").asText());
            for (String field : List.of("merges", "conditionalFormats", "dataValidations", "dataRegions", "sheetTables", "spillRanges", "protectionRules", "sparklines")) {
                JsonNode values = sheet.get(field);
                if (values != null && values.isArray()) for (JsonNode value : values) if (containsDeletedRange(value, sourceSheetId)) references.add(field + ":" + value.path("id").asText());
            }
            for (String field : List.of("conditionalFormats", "dataValidations")) {
                JsonNode values = sheet.get(field);
                if (values != null && values.isArray()) for (JsonNode value : values) inspectRuleReferences(references, value, sourceName, sourceSheetId, field);
            }
            for (String field : List.of("bandedRule", "autoFilter")) {
                JsonNode value = sheet.get(field);
                if (containsDeletedRange(value, sourceSheetId)) references.add(field + ":" + value.path("id").asText());
            }
            JsonNode pivots = sheet.get("pivots");
            if (pivots != null && pivots.isArray()) for (JsonNode pivot : pivots) if (containsDeletedPivotReference(pivot, sourceSheetId, deletedSheetTableIds)) references.add("pivot:" + pivot.path("id").asText());
            JsonNode drawings = sheet.get("drawings");
            if (drawings != null && drawings.isArray()) for (JsonNode drawing : drawings) {
                if (sourceSheetId.equals(drawing.path("sheetId").asText())) references.add("drawing-owner:" + drawing.path("id").asText());
            }
            JsonNode drawingGroups = sheet.get("drawingGroups");
            if (drawingGroups != null && drawingGroups.isArray()) for (JsonNode group : drawingGroups) {
                if (sourceSheetId.equals(group.path("sheetId").asText())) references.add("drawing-group-owner:" + group.path("id").asText());
                JsonNode members = group.get("memberDrawingIds");
                if (members != null && members.isArray()) for (JsonNode member : members) {
                    if (deletedDrawingIds.contains(member.asText())) references.add("drawing-group-member:" + group.path("id").asText());
                }
            }
            inspectDrawingPayloadReferences(references, sheet.get("drawingPayloads"), sourceSheetId, sourceName, deletedPivotIds, deletedSheetTableIds, deletedDrawingIds);
            if (sourceSheetId.equals(sheet.path("reportSheet").path("templateSheetId").asText())) references.add("report-template:" + sheet.path("id").asText());
        }
        JsonNode names = root.get("definedNameModels");
        if (names != null && names.isArray()) for (JsonNode name : names) {
            if ("sheet".equals(name.path("scope").asText()) && sourceSheetId.equals(name.path("sheetId").asText())) continue;
            addFormulaReference(references, "defined-name:" + name.path("name").asText(), name.get("formula"), sourceName);
            if (sourceSheetId.equals(name.path("anchor").path("sheetId").asText())) references.add("defined-name-anchor:" + name.path("name").asText());
        }
        JsonNode legacyNames = root.get("definedNames");
        if (legacyNames != null && legacyNames.isObject()) legacyNames.fields().forEachRemaining(entry ->
                addFormulaReference(references, "defined-name-projection:" + entry.getKey(), entry.getValue(), sourceName));
        inspectWorkbookDataModelReferences(references, root, sourceSheetId, sourceName);
        inspectWorkbookStyleTemplateReferences(references, root, sourceSheetId, sourceName);
        inspectQueryReferences(references, root, sourceSheetId, deletedPivotIds, deletedSheetTableIds);
        inspectPrintDocumentReferences(references, root, sourceSheetId);
        if (!references.isEmpty()) throw ServiceException.conflict("Cannot delete sheet with external references: " + String.join(", ", references));
    }

    private Set<String> collectIds(JsonNode values) {
        Set<String> ids = new HashSet<>();
        if (values == null || values.isNull()) return ids;
        if (!values.isArray()) throw ServiceException.validation("Worksheet reference owner must be an array");
        for (JsonNode value : values) if (value.path("id").isTextual()) ids.add(value.path("id").asText());
        return ids;
    }

    private void addFormulaReference(List<String> references, String participant, JsonNode formula, String sourceName) {
        if (formula == null || !formula.isTextual()) return;
        String value = formula.asText();
        if (!value.equals(FormulaReferenceTransformer.renameSheet(value, sourceName, sourceName + "__deleted__"))) {
            references.add(participant + ":" + value);
        }
    }

    private void inspectRuleReferences(List<String> references, JsonNode rule, String sourceName, String sourceSheetId, String ownerKind) {
        if (!rule.isObject()) return;
        String ruleId = rule.path("id").asText(ownerKind);
        if (sourceSheetId.equals(rule.path("formulaAnchor").path("sheetId").asText())) references.add(ownerKind + "-formula-anchor:" + ruleId);
        JsonNode listSource = rule.get("listSource");
        if (listSource != null && sourceSheetId.equals(listSource.path("range").path("sheetId").asText())) references.add(ownerKind + "-list-range:" + ruleId);
        for (String field : List.of("value1", "value2", "formula1", "formula2")) {
            addFormulaReference(references, ownerKind + "." + field + ":" + ruleId, rule.get(field), sourceName);
        }
        if (listSource != null && "formula".equals(listSource.path("kind").asText())) {
            addFormulaReference(references, ownerKind + ".listSource:" + ruleId, listSource.get("formula"), sourceName);
        }
    }

    private void inspectDrawingPayloadReferences(
            List<String> references,
            JsonNode payloads,
            String sourceSheetId,
            String sourceName,
            Set<String> deletedPivotIds,
            Set<String> deletedSheetTableIds,
            Set<String> deletedDrawingIds
    ) {
        if (payloads == null || payloads.isNull()) return;
        if (!payloads.isObject()) throw ServiceException.validation("drawingPayloads must be an object");
        payloads.fields().forEachRemaining(entry -> {
            JsonNode payload = entry.getValue();
            if (!payload.isObject()) throw ServiceException.validation("Drawing payload must be an object: " + entry.getKey());
            String participant = "drawing-payload:" + entry.getKey();
            String kind = payload.path("kind").asText();
            if ("camera".equals(kind) || "screenshot".equals(kind)) {
                if (sourceSheetId.equals(payload.path("sourceRange").path("sheetId").asText())) references.add(participant + ":sourceRange");
            } else if ("form-control".equals(kind)) {
                if (sourceSheetId.equals(payload.path("cellLink").path("sheetId").asText())) references.add(participant + ":cellLink");
                if (sourceSheetId.equals(payload.path("inputRange").path("sheetId").asText())) references.add(participant + ":inputRange");
            } else if ("connector".equals(kind)) {
                for (String endpoint : List.of("start", "end")) {
                    if (deletedDrawingIds.contains(payload.path(endpoint).path("drawingId").asText())) references.add(participant + ":" + endpoint);
                }
            } else if ("shape".equals(kind)) {
                JsonNode hyperlink = payload.get("hyperlink");
                if (hyperlink != null && "sheet".equals(hyperlink.path("kind").asText())
                        && sourceSheetId.equals(hyperlink.path("sheetId").asText())) references.add(participant + ":hyperlink");
                addFormulaReference(references, participant + ".propertyFormula", payload.get("propertyFormula"), sourceName);
            } else if ("chart".equals(kind)) {
                for (String field : CHART_TEXT_FORMULA_FIELDS) {
                    addFormulaReference(references, participant + "." + field,
                            readChartTextFormula((ObjectNode) payload, field), sourceName);
                }
                JsonNode source = payload.get("source");
                if (source == null || !source.isObject()) throw ServiceException.validation("Chart source must be canonical");
                String sourceKind = source.path("kind").asText();
                if ("worksheet-ranges".equals(sourceKind)) {
                    JsonNode ranges = optionalArray(source, "ranges");
                    if (ranges != null) for (JsonNode range : ranges) if (sourceSheetId.equals(range.path("sheetId").asText())) references.add(participant + ":source");
                } else if ("report-range".equals(sourceKind)) {
                    if (sourceSheetId.equals(source.path("range").path("sheetId").asText())) references.add(participant + ":source");
                } else if ("pivot".equals(sourceKind) && deletedPivotIds.contains(source.path("pivotId").asText())) {
                    references.add(participant + ":pivot-source");
                } else if ("table".equals(sourceKind) && deletedSheetTableIds.contains(source.path("tableId").asText())) {
                    references.add(participant + ":table-source");
                }
                if (sourceSheetId.equals(payload.path("categoryRange").path("sheetId").asText())) references.add(participant + ":categoryRange");
                JsonNode seriesValues = optionalArray(payload, "series");
                if (seriesValues != null) for (JsonNode series : seriesValues) {
                    for (String field : List.of("range", "xRange", "yRange", "sizeRange", "categoryRange")) {
                        if (sourceSheetId.equals(series.path(field).path("sheetId").asText())) references.add(participant + ".series." + field);
                    }
                    JsonNode stockRoles = series.get("stockRoles");
                    if (stockRoles != null && stockRoles.isObject()) for (String field : List.of("open", "high", "low", "close", "volume")) {
                        if (sourceSheetId.equals(stockRoles.path(field).path("sheetId").asText())) references.add(participant + ".series.stockRoles." + field);
                    }
                    JsonNode errorBars = series.get("errorBars");
                    if (errorBars != null && errorBars.isObject()) for (String field : List.of("plusRange", "minusRange")) {
                        if (sourceSheetId.equals(errorBars.path(field).path("sheetId").asText())) references.add(participant + ".series.errorBars." + field);
                    }
                    JsonNode dataLabels = series.get("dataLabels");
                    if (dataLabels != null && sourceSheetId.equals(dataLabels.path("valuesFromCells").path("sheetId").asText())) references.add(participant + ".series.dataLabels.valuesFromCells");
                }
            } else if ("slicer".equals(kind) || "timeline".equals(kind)) {
                if (deletedPivotIds.contains(payload.path("pivotId").asText())) references.add(participant + ":pivot");
                JsonNode connections = optionalArray(payload, "connections");
                if (connections != null) for (JsonNode connection : connections) {
                    if (deletedPivotIds.contains(connection.path("pivotId").asText())) references.add(participant + ":connection");
                }
            }
        });
    }

    private void inspectWorkbookDataModelReferences(List<String> references, ObjectNode root, String sourceSheetId, String sourceName) {
        JsonNode dataModel = root.get("dataModel");
        if (dataModel == null || dataModel.isNull()) return;
        if (!dataModel.isObject()) throw ServiceException.validation("dataModel must be an object");
        for (String collection : List.of("sources", "tables")) {
            JsonNode values = optionalArray(dataModel, collection);
            if (values == null) continue;
            for (JsonNode value : values) {
                if (sourceSheetId.equals(value.path("sourceSheetId").asText())
                        || sourceSheetId.equals(value.path("sourceRange").path("sheetId").asText())) references.add("workbook-" + collection + ":" + value.path("id").asText());
            }
        }
        JsonNode views = optionalArray(dataModel, "views");
        if (views != null) for (JsonNode view : views) {
            JsonNode fields = optionalArray(view, "fields");
            if (fields == null) continue;
            for (JsonNode field : fields) addFormulaReference(references, "data-view:" + view.path("id").asText() + "." + field.path("fieldId").asText(), field.get("formula"), sourceName);
        }
    }

    private void inspectWorkbookStyleTemplateReferences(List<String> references, ObjectNode root, String sourceSheetId, String sourceName) {
        JsonNode templates = optionalArray(root, "cellStyleTemplates");
        if (templates == null) return;
        for (JsonNode template : templates) {
            JsonNode validation = template.get("dataValidation");
            if (validation == null || !validation.isObject()) continue;
            String templateId = template.path("id").asText();
            if (sourceSheetId.equals(validation.path("formulaAnchor").path("sheetId").asText())) references.add("cell-style-template-anchor:" + templateId);
            JsonNode listSource = validation.get("listSource");
            if (listSource != null && sourceSheetId.equals(listSource.path("range").path("sheetId").asText())) references.add("cell-style-template-list-range:" + templateId);
            for (String field : List.of("formula1", "formula2")) addFormulaReference(references, "cell-style-template." + field + ":" + templateId, validation.get(field), sourceName);
            if (listSource != null && "formula".equals(listSource.path("kind").asText())) {
                addFormulaReference(references, "cell-style-template.listSource:" + templateId, listSource.get("formula"), sourceName);
            }
        }
    }

    private void inspectQueryReferences(List<String> references, ObjectNode root, String sourceSheetId, Set<String> deletedPivotIds, Set<String> deletedSheetTableIds) {
        JsonNode queries = optionalArray(root, "queryDefinitions");
        if (queries == null) return;
        for (JsonNode query : queries) {
            JsonNode target = query.get("lastTarget");
            if (target == null || !target.isObject()) continue;
            String queryId = query.path("id").asText();
            if (sourceSheetId.equals(target.path("sheetId").asText())) references.add("query-load-target:" + queryId);
            if ("pivot-source".equals(target.path("kind").asText()) && deletedPivotIds.contains(target.path("pivotId").asText())) references.add("query-pivot-target:" + queryId);
            if ("sheet-table".equals(target.path("kind").asText()) && deletedSheetTableIds.contains(target.path("tableId").asText())) references.add("query-sheet-table-target:" + queryId);
        }
    }

    private void inspectPrintDocumentReferences(List<String> references, ObjectNode root, String sourceSheetId) {
        JsonNode documents = optionalArray(root, "printDocuments");
        if (documents == null) return;
        for (JsonNode document : documents) {
            if (sourceSheetId.equals(document.path("sheetId").asText())) continue;
            String ownerId = document.path("sheetId").asText();
            JsonNode areas = optionalArray(document, "printAreas");
            if (areas != null) for (JsonNode area : areas) {
                if (sourceSheetId.equals(area.path("sheetId").asText()) || sourceSheetId.equals(area.path("range").path("sheetId").asText())) references.add("print-document-area:" + ownerId);
            }
            JsonNode breaks = optionalArray(document, "pageBreaks");
            if (breaks != null) for (JsonNode pageBreak : breaks) {
                if (sourceSheetId.equals(pageBreak.path("sheetId").asText())) references.add("print-document-page-break:" + ownerId);
            }
        }
    }

    private boolean containsDeletedRange(JsonNode value, String sourceSheetId) {
        if (value == null || value.isNull() || !value.isObject()) return false;
        if (sourceSheetId.equals(value.path("sheetId").asText())) return true;
        if (sourceSheetId.equals(value.path("range").path("sheetId").asText())) return true;
        if (sourceSheetId.equals(value.path("sourceRange").path("sheetId").asText())) return true;
        if (sourceSheetId.equals(value.path("formulaAnchor").path("sheetId").asText())) return true;
        JsonNode ranges = value.get("ranges");
        if (ranges != null && ranges.isArray()) for (JsonNode range : ranges) if (sourceSheetId.equals(range.path("sheetId").asText())) return true;
        JsonNode listSource = value.path("listSource");
        return sourceSheetId.equals(listSource.path("range").path("sheetId").asText());
    }

    private boolean containsDeletedPivotReference(JsonNode value, String sourceSheetId, Set<String> deletedSheetTableIds) {
        if (!value.isObject()) return false;
        if (sourceSheetId.equals(value.path("target").path("sheetId").asText())) return true;
        JsonNode source = value.get("source");
        if (source == null || !source.isObject()) return false;
        if (deletedSheetTableIds.contains(source.path("tableId").asText())) return true;
        if (sourceSheetId.equals(source.path("range").path("sheetId").asText()) || sourceSheetId.equals(source.path("sheetId").asText())) return true;
        JsonNode ranges = source.get("ranges");
        if (ranges != null && ranges.isArray()) for (JsonNode range : ranges) if (sourceSheetId.equals(range.path("range").path("sheetId").asText())) return true;
        return false;
    }

    private void removeMatching(ArrayNode values, java.util.function.Predicate<JsonNode> predicate) {
        for (int index = values.size() - 1; index >= 0; index--) {
            if (predicate.test(values.get(index))) values.remove(index);
        }
    }

    /**
     * Explicit participant registry for duplicate. A workbook snapshot may
     * contain arbitrary preserved JSON, so a field-name recursive rewrite is
     * intentionally forbidden here. Every supported owner is mapped below.
     */
    private void remapDuplicatedSheetReferences(
            ObjectNode root,
            ObjectNode copy,
            String sourceSheetId,
            String sourceName,
            String targetSheetId,
            String targetName
    ) {
        Map<String, String> pivotIds = remapIds(root, copy, "pivots", targetSheetId);
        Map<String, String> sparklineIds = remapIds(root, copy, "sparklines", targetSheetId);
        Map<String, String> sparklineGroupIds = remapIds(root, copy, "sparklineGroups", targetSheetId);
        Map<String, String> drawingIds = remapIds(root, copy, "drawings", targetSheetId);
        Map<String, String> drawingGroupIds = remapIds(root, copy, "drawingGroups", targetSheetId);
        Map<String, String> conditionalFormatIds = remapIds(root, copy, "conditionalFormats", targetSheetId);
        Map<String, String> dataValidationIds = remapIds(root, copy, "dataValidations", targetSheetId);
        Map<String, String> tableIds = remapIds(root, copy, "sheetTables", targetSheetId);
        Map<String, String> payloadIds = remapObjectKeys(root, copy, "drawingPayloads", targetSheetId);

        remapRangeArray(copy, "merges", sourceSheetId, targetSheetId, null);
        remapRangeArray(copy, "conditionalFormats", sourceSheetId, targetSheetId, conditionalFormatIds);
        remapRangeArray(copy, "dataValidations", sourceSheetId, targetSheetId, dataValidationIds);
        remapRangeArray(copy, "spillRanges", sourceSheetId, targetSheetId, null);
        remapRangeArray(copy, "protectionRules", sourceSheetId, targetSheetId, null);
        remapRangeArray(copy, "dataRegions", sourceSheetId, targetSheetId, null);
        remapRangeObject(copy, "bandedRule", sourceSheetId, targetSheetId);
        remapAutoFilter(copy, sourceSheetId, targetSheetId);
        remapPivots(copy, sourceSheetId, targetSheetId, pivotIds, tableIds);
        remapSparklines(copy, sourceSheetId, targetSheetId, sparklineIds, sparklineGroupIds);
        remapSparklineGroups(copy, sourceSheetId, targetSheetId, sparklineGroupIds, sparklineIds);
        remapSheetTables(copy, sourceSheetId, targetSheetId, tableIds);
        remapDrawings(copy, sourceSheetId, targetSheetId, drawingIds, payloadIds);
        remapDrawingPayloads(copy, sourceSheetId, targetSheetId, drawingIds, pivotIds, tableIds);
        remapDrawingGroups(copy, sourceSheetId, targetSheetId, drawingGroupIds, drawingIds);
        remapHyperlinks(copy, sourceSheetId, targetSheetId);
        remapReview(root, copy, targetSheetId);
        remapReportSheet(copy, sourceSheetId, targetSheetId, tableIds);
        remapCopiedFormulas(copy, sourceName, targetName);
    }

    private Map<String, String> remapIds(ObjectNode root, ObjectNode copy, String arrayField, String targetSheetId) {
        Map<String, String> result = new HashMap<>();
        Set<String> existing = collectIds(root, arrayField);
        JsonNode values = copy.get(arrayField);
        if (values == null || !values.isArray()) return result;
        for (JsonNode raw : values) {
            if (!raw.isObject() || !raw.path("id").isTextual()) continue;
            String oldId = raw.path("id").asText();
            String newId = allocateId(existing, oldId, targetSheetId);
            result.put(oldId, newId);
            existing.add(newId);
            ((ObjectNode) raw).put("id", newId);
        }
        return result;
    }

    private Map<String, String> remapObjectKeys(ObjectNode root, ObjectNode copy, String field, String targetSheetId) {
        Map<String, String> result = new HashMap<>();
        Set<String> existing = new HashSet<>();
        for (JsonNode sheet : SnapshotMutationSupport.sheets(root)) {
            JsonNode values = sheet.get(field);
            if (values != null && values.isObject()) values.fieldNames().forEachRemaining(existing::add);
        }
        JsonNode values = copy.get(field);
        if (values == null || !values.isObject()) return result;
        ObjectNode remapped = JsonNodeFactory.instance.objectNode();
        values.fields().forEachRemaining(entry -> {
            String newId = allocateId(existing, entry.getKey(), targetSheetId);
            existing.add(newId);
            result.put(entry.getKey(), newId);
            remapped.set(newId, entry.getValue());
        });
        copy.set(field, remapped);
        return result;
    }

    private Set<String> collectIds(ObjectNode root, String arrayField) {
        Set<String> ids = new HashSet<>();
        for (JsonNode sheet : SnapshotMutationSupport.sheets(root)) {
            JsonNode values = sheet.get(arrayField);
            if (values != null && values.isArray()) for (JsonNode value : values) if (value.path("id").isTextual()) ids.add(value.path("id").asText());
        }
        return ids;
    }

    private String allocateId(Set<String> existing, String sourceId, String targetSheetId) {
        String stem = sourceId + "::" + targetSheetId;
        String candidate = stem;
        int suffix = 2;
        while (existing.contains(candidate)) candidate = stem + "::" + suffix++;
        return candidate;
    }

    private void remapRange(JsonNode value, String sourceSheetId, String targetSheetId) {
        if (value != null && value.isObject() && sourceSheetId.equals(value.path("sheetId").asText())) ((ObjectNode) value).put("sheetId", targetSheetId);
    }

    private void remapRangeObject(ObjectNode owner, String field, String sourceSheetId, String targetSheetId) {
        remapRange(owner.get(field), sourceSheetId, targetSheetId);
    }

    private void remapRangeArray(ObjectNode owner, String field, String sourceSheetId, String targetSheetId, Map<String, String> ids) {
        JsonNode values = owner.get(field);
        if (values == null || !values.isArray()) return;
        for (JsonNode value : values) {
            if (!value.isObject()) continue;
            ObjectNode object = (ObjectNode) value;
            if (ids != null && object.path("id").isTextual()) object.put("id", ids.getOrDefault(object.path("id").asText(), object.path("id").asText()));
            if (object.has("sheetId")) object.put("sheetId", targetSheetId);
            remapRange(object.get("range"), sourceSheetId, targetSheetId);
            remapRange(object.get("formulaAnchor"), sourceSheetId, targetSheetId);
            remapRange(object.get("sourceRange"), sourceSheetId, targetSheetId);
            remapRange(object.get("affectedBand"), sourceSheetId, targetSheetId);
            JsonNode ranges = object.get("ranges");
            if (ranges != null && ranges.isArray()) for (JsonNode range : ranges) remapRange(range, sourceSheetId, targetSheetId);
            JsonNode listSource = object.get("listSource");
            if (listSource != null && listSource.isObject()) remapRange(listSource.get("range"), sourceSheetId, targetSheetId);
        }
    }

    private void remapAutoFilter(ObjectNode copy, String sourceSheetId, String targetSheetId) {
        JsonNode filter = copy.get("autoFilter");
        if (filter != null && filter.isObject()) {
            ((ObjectNode) filter).put("sheetId", targetSheetId);
            remapRange(filter.get("range"), sourceSheetId, targetSheetId);
        }
    }

    private void remapPivots(ObjectNode copy, String sourceSheetId, String targetSheetId, Map<String, String> pivotIds, Map<String, String> tableIds) {
        JsonNode values = copy.get("pivots");
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (!raw.isObject()) continue;
            ObjectNode pivot = (ObjectNode) raw;
            if (pivot.path("target").isObject()) remapRange(pivot.get("target"), sourceSheetId, targetSheetId);
            JsonNode source = pivot.get("source");
            if (source != null && source.isObject()) {
                ObjectNode sourceObject = (ObjectNode) source;
                if (sourceObject.path("kind").asText().equals("table") && sourceObject.path("tableId").isTextual()) sourceObject.put("tableId", tableIds.getOrDefault(sourceObject.path("tableId").asText(), sourceObject.path("tableId").asText()));
                remapRange(sourceObject.get("range"), sourceSheetId, targetSheetId);
                JsonNode ranges = sourceObject.get("ranges");
                if (ranges != null && ranges.isArray()) for (JsonNode entry : ranges) if (entry.isObject()) remapRange(entry.get("range"), sourceSheetId, targetSheetId);
                remapRange(sourceObject, sourceSheetId, targetSheetId);
            }
        }
    }

    private void remapSparklines(ObjectNode copy, String sourceSheetId, String targetSheetId, Map<String, String> sparklineIds, Map<String, String> groupIds) {
        JsonNode values = copy.get("sparklines");
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (!raw.isObject()) continue;
            ObjectNode sparkline = (ObjectNode) raw;
            sparkline.put("sheetId", targetSheetId);
            remapRange(sparkline.get("sourceRange"), sourceSheetId, targetSheetId);
            if (sparkline.path("groupId").isTextual()) sparkline.put("groupId", groupIds.getOrDefault(sparkline.path("groupId").asText(), sparkline.path("groupId").asText()));
        }
    }

    private void remapSparklineGroups(ObjectNode copy, String sourceSheetId, String targetSheetId, Map<String, String> groupIds, Map<String, String> sparklineIds) {
        JsonNode values = copy.get("sparklineGroups");
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (!raw.isObject()) continue;
            ObjectNode group = (ObjectNode) raw;
            group.put("sheetId", targetSheetId);
            if (group.path("id").isTextual()) group.put("id", groupIds.getOrDefault(group.path("id").asText(), group.path("id").asText()));
            JsonNode ids = group.get("sparklineIds");
            if (ids != null && ids.isArray()) for (int index = 0; index < ids.size(); index++) if (ids.get(index).isTextual()) ((ArrayNode) ids).set(index, JsonNodeFactory.instance.textNode(sparklineIds.getOrDefault(ids.get(index).asText(), ids.get(index).asText())));
        }
    }

    private void remapSheetTables(ObjectNode copy, String sourceSheetId, String targetSheetId, Map<String, String> tableIds) {
        JsonNode values = copy.get("sheetTables");
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (!raw.isObject()) continue;
            ObjectNode table = (ObjectNode) raw;
            table.put("sheetId", targetSheetId);
            remapRange(table.get("range"), sourceSheetId, targetSheetId);
            if (table.path("autoFilter").isObject()) {
                ObjectNode filter = (ObjectNode) table.get("autoFilter");
                filter.put("sheetId", targetSheetId);
                remapRange(filter.get("range"), sourceSheetId, targetSheetId);
            }
        }
    }

    private void remapDrawings(ObjectNode copy, String sourceSheetId, String targetSheetId, Map<String, String> drawingIds, Map<String, String> payloadIds) {
        JsonNode values = copy.get("drawings");
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (!raw.isObject()) continue;
            ObjectNode drawing = (ObjectNode) raw;
            String oldId = drawing.path("id").asText();
            drawing.put("id", drawingIds.getOrDefault(oldId, oldId));
            drawing.put("sheetId", targetSheetId);
            if (drawing.path("payloadId").isTextual()) drawing.put("payloadId", payloadIds.getOrDefault(drawing.path("payloadId").asText(), drawing.path("payloadId").asText()));
        }
    }

    private void remapDrawingPayloads(ObjectNode copy, String sourceSheetId, String targetSheetId, Map<String, String> drawingIds, Map<String, String> pivotIds, Map<String, String> tableIds) {
        JsonNode values = copy.get("drawingPayloads");
        if (values == null || !values.isObject()) return;
        values.fields().forEachRemaining(entry -> remapPayload(entry.getValue(), sourceSheetId, targetSheetId, drawingIds, pivotIds, tableIds));
    }

    private void remapPayload(JsonNode raw, String sourceSheetId, String targetSheetId, Map<String, String> drawingIds, Map<String, String> pivotIds, Map<String, String> tableIds) {
        if (raw == null || !raw.isObject()) return;
        ObjectNode payload = (ObjectNode) raw;
        String kind = payload.path("kind").asText();
        if ("camera".equals(kind) || "screenshot".equals(kind)) remapRange(payload.get("sourceRange"), sourceSheetId, targetSheetId);
        if ("chart".equals(kind)) {
            remapRange(payload.get("categoryRange"), sourceSheetId, targetSheetId);
            JsonNode source = payload.get("source");
            if (source == null || !source.isObject()) throw ServiceException.validation("Chart source must be canonical");
            String sourceKind = source.path("kind").asText();
            if ("worksheet-ranges".equals(sourceKind)) {
                JsonNode ranges = source.get("ranges");
                if (ranges == null || !ranges.isArray()) throw ServiceException.validation("Chart worksheet source ranges are invalid");
                for (JsonNode range : ranges) remapRange(range, sourceSheetId, targetSheetId);
            } else if ("report-range".equals(sourceKind)) {
                remapRange(source.get("range"), sourceSheetId, targetSheetId);
            } else if ("pivot".equals(sourceKind)) {
                if (source.path("pivotId").isTextual()) ((ObjectNode) source).put("pivotId", pivotIds.getOrDefault(source.path("pivotId").asText(), source.path("pivotId").asText()));
            } else if (!"table".equals(sourceKind)) {
                throw ServiceException.validation("Chart source kind is invalid: " + sourceKind);
            }
            for (JsonNode series : payload.path("series")) {
                if (!series.isObject()) throw ServiceException.validation("Chart series is invalid");
                for (String field : List.of("range", "xRange", "yRange", "sizeRange", "categoryRange")) remapRange(series.get(field), sourceSheetId, targetSheetId);
                JsonNode errorBars = series.get("errorBars");
                if (errorBars != null && errorBars.isObject()) {
                    remapRange(errorBars.get("plusRange"), sourceSheetId, targetSheetId);
                    remapRange(errorBars.get("minusRange"), sourceSheetId, targetSheetId);
                }
                JsonNode stockRoles = series.get("stockRoles");
                if (stockRoles != null && stockRoles.isObject()) for (String field : List.of("open", "high", "low", "close", "volume")) remapRange(stockRoles.get(field), sourceSheetId, targetSheetId);
                JsonNode dataLabels = series.get("dataLabels");
                if (dataLabels != null && dataLabels.isObject()) remapRange(dataLabels.get("valuesFromCells"), sourceSheetId, targetSheetId);
            }
        }
        if ("connector".equals(kind)) {
            if (payload.path("start").path("drawingId").isTextual()) ((ObjectNode) payload.get("start")).put("drawingId", drawingIds.getOrDefault(payload.path("start").path("drawingId").asText(), payload.path("start").path("drawingId").asText()));
            if (payload.path("end").path("drawingId").isTextual()) ((ObjectNode) payload.get("end")).put("drawingId", drawingIds.getOrDefault(payload.path("end").path("drawingId").asText(), payload.path("end").path("drawingId").asText()));
        }
        if ("form-control".equals(kind)) {
            remapRange(payload.get("cellLink"), sourceSheetId, targetSheetId);
            remapRange(payload.get("inputRange"), sourceSheetId, targetSheetId);
        }
        if ("slicer".equals(kind) || "timeline".equals(kind)) {
            if (payload.path("pivotId").isTextual()) payload.put("pivotId", pivotIds.getOrDefault(payload.path("pivotId").asText(), payload.path("pivotId").asText()));
            JsonNode connections = payload.get("connections");
            if (connections != null && connections.isArray()) for (JsonNode connection : connections) {
                if (connection.isObject() && connection.path("pivotId").isTextual()) {
                    ((ObjectNode) connection).put("pivotId", pivotIds.getOrDefault(connection.path("pivotId").asText(), connection.path("pivotId").asText()));
                }
            }
        }
    }

    private void remapDrawingGroups(ObjectNode copy, String sourceSheetId, String targetSheetId, Map<String, String> groupIds, Map<String, String> drawingIds) {
        JsonNode values = copy.get("drawingGroups");
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (!raw.isObject()) continue;
            ObjectNode group = (ObjectNode) raw;
            group.put("sheetId", targetSheetId);
            if (group.path("id").isTextual()) group.put("id", groupIds.getOrDefault(group.path("id").asText(), group.path("id").asText()));
            JsonNode members = group.get("memberDrawingIds");
            if (members != null && members.isArray()) for (int index = 0; index < members.size(); index++) if (members.get(index).isTextual()) ((ArrayNode) members).set(index, JsonNodeFactory.instance.textNode(drawingIds.getOrDefault(members.get(index).asText(), members.get(index).asText())));
        }
    }

    private void remapHyperlinks(ObjectNode copy, String sourceSheetId, String targetSheetId) {
        JsonNode values = copy.get("hyperlinks");
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (raw.isObject() && raw.path("hyperlink").path("target").path("sheetId").asText().equals(sourceSheetId)) ((ObjectNode) raw.path("hyperlink").path("target")).put("sheetId", targetSheetId);
        }
    }

    private void remapReview(ObjectNode root, ObjectNode copy, String targetSheetId) {
        ObjectNode review = SnapshotMutationSupport.review(copy);
        Set<String> existing = new HashSet<>();
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject() || !rawSheet.path("review").isObject()) continue;
            ObjectNode current = (ObjectNode) rawSheet.get("review");
            current.path("notesById").fieldNames().forEachRemaining(existing::add);
            current.path("threadsById").fields().forEachRemaining(entry -> {
                existing.add(entry.getKey());
                for (JsonNode reply : entry.getValue().path("replies")) if (reply.path("id").isTextual()) existing.add(reply.path("id").asText());
            });
        }
        Map<String, String> ids = new HashMap<>();
        ObjectNode notesById = SnapshotMutationSupport.requiredObject(review, "notesById");
        ObjectNode remappedNotes = notesById.objectNode();
        notesById.fields().forEachRemaining(entry -> {
            String nextId = allocateId(existing, entry.getKey(), targetSheetId);
            existing.add(nextId);
            ids.put(entry.getKey(), nextId);
            ObjectNode note = (ObjectNode) entry.getValue().deepCopy();
            note.put("id", nextId);
            remappedNotes.set(nextId, note);
        });
        review.set("notesById", remappedNotes);
        ObjectNode threadsById = SnapshotMutationSupport.requiredObject(review, "threadsById");
        ObjectNode remappedThreads = threadsById.objectNode();
        threadsById.fields().forEachRemaining(entry -> {
            String nextId = allocateId(existing, entry.getKey(), targetSheetId);
            existing.add(nextId);
            ids.put(entry.getKey(), nextId);
            ObjectNode thread = (ObjectNode) entry.getValue().deepCopy();
            thread.put("id", nextId);
            thread.put("sheetId", targetSheetId);
            ArrayNode replies = (ArrayNode) thread.path("replies").deepCopy();
            for (JsonNode reply : replies) if (reply.isObject()) {
                String replyId = reply.path("id").asText(null);
                if (replyId != null && !replyId.isBlank()) {
                    String nextReplyId = allocateId(existing, replyId, targetSheetId);
                    existing.add(nextReplyId);
                    ids.put(replyId, nextReplyId);
                    ((ObjectNode) reply).put("id", nextReplyId);
                }
            }
            thread.set("replies", replies);
            remappedThreads.set(nextId, thread);
        });
        review.set("threadsById", remappedThreads);
        ObjectNode notesByCell = SnapshotMutationSupport.requiredObject(review, "notesByCell");
        notesByCell.fields().forEachRemaining(entry -> entry.setValue(JsonNodeFactory.instance.textNode(ids.getOrDefault(entry.getValue().asText(), entry.getValue().asText()))));
        ObjectNode threadIdsByCell = SnapshotMutationSupport.requiredObject(review, "threadIdsByCell");
        threadIdsByCell.fields().forEachRemaining(entry -> {
            if (!entry.getValue().isArray()) throw ServiceException.validation("Review thread cell index is invalid");
            ArrayNode values = (ArrayNode) entry.getValue();
            for (int index = 0; index < values.size(); index++) values.set(index, JsonNodeFactory.instance.textNode(ids.getOrDefault(values.get(index).asText(), values.get(index).asText())));
        });
    }

    private void remapReportSheet(ObjectNode copy, String sourceSheetId, String targetSheetId, Map<String, String> tableIds) {
        JsonNode report = copy.get("reportSheet");
        if (report != null && report.isObject()) {
            ObjectNode object = (ObjectNode) report;
            if (object.path("templateSheetId").asText().equals(sourceSheetId)) object.put("templateSheetId", targetSheetId);
            if (object.path("tableId").isTextual()) object.put("tableId", tableIds.getOrDefault(object.path("tableId").asText(), object.path("tableId").asText()));
        }
    }

    private void remapCopiedFormulas(ObjectNode copy, String sourceName, String targetName) {
        JsonNode cells = copy.get("cells");
        if (cells != null && cells.isObject()) cells.fields().forEachRemaining(row -> {
            if (!row.getValue().isObject()) return;
            row.getValue().fields().forEachRemaining(cell -> {
                if (!cell.getValue().isObject()) return;
                ObjectNode owner = (ObjectNode) cell.getValue();
                rewriteFormulaField(owner, "formula", sourceName, targetName);
                JsonNode metadata = owner.get("formulaMetadata");
                if (metadata != null && metadata.isObject()) rewriteFormulaMetadataSource((ObjectNode) metadata, sourceName, targetName, "duplicated");
                JsonNode presentation = owner.get("presentation");
                JsonNode barcodeSource = presentation == null ? null : presentation.get("source");
                if (presentation != null && "barcode".equals(presentation.path("kind").asText())
                        && barcodeSource != null && "formula".equals(barcodeSource.path("kind").asText())) {
                    rewriteFormulaField((ObjectNode) barcodeSource, "formula", sourceName, targetName);
                }
            });
        });
        JsonNode tableSheet = copy.get("tableSheet");
        if (tableSheet != null && tableSheet.isObject()) {
            JsonNode columns = optionalArray(tableSheet, "columns");
            if (columns != null) for (JsonNode column : columns) if (column.isObject()) rewriteFormulaField((ObjectNode) column, "formula", sourceName, targetName);
        }
        JsonNode payloads = copy.get("drawingPayloads");
        if (payloads != null && payloads.isObject()) payloads.fields().forEachRemaining(entry -> {
            JsonNode payload = entry.getValue();
            if (payload.isObject() && "shape".equals(payload.path("kind").asText())) {
                rewriteFormulaField((ObjectNode) payload, "propertyFormula", sourceName, targetName);
            } else if (payload.isObject() && "chart".equals(payload.path("kind").asText())) {
                rewriteChartTextFormulas((ObjectNode) payload, sourceName, targetName);
            }
        });
        rewriteRuleFormulaFields(copy.get("conditionalFormats"), sourceName, targetName);
        rewriteRuleFormulaFields(copy.get("dataValidations"), sourceName, targetName);
    }

    private void rewriteRuleFormulaFields(JsonNode values, String previousName, String nextName) {
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (!raw.isObject()) continue;
            ObjectNode rule = (ObjectNode) raw;
            for (String field : List.of("value1", "value2", "formula1", "formula2")) rewriteFormulaField(rule, field, previousName, nextName);
            JsonNode listSource = rule.get("listSource");
            if (listSource != null && listSource.isObject() && listSource.path("kind").asText().equals("formula")) rewriteFormulaField((ObjectNode) listSource, "formula", previousName, nextName);
        }
    }

    private void rewriteFormulaField(ObjectNode owner, String field, String previousName, String nextName) {
        JsonNode formula = owner.get(field);
        if (formula != null && formula.isTextual()) owner.put(field, renameFormula(formula.asText(), previousName, nextName));
    }

    private void rewriteChartTextFormulas(ObjectNode payload, String previousName, String nextName) {
        for (String field : CHART_TEXT_FORMULA_FIELDS) {
            JsonNode formula = readChartTextFormula(payload, field);
            if (formula != null) {
                chartTextFormulaModel(payload, field).put("linkedFormula", renameFormula(formula.asText(), previousName, nextName));
            }
        }
    }

    private JsonNode readChartTextFormula(ObjectNode payload, String field) {
        JsonNode formula = chartTextFormulaModel(payload, field).get("linkedFormula");
        if (formula == null || formula.isNull()) return null;
        if (!formula.isTextual() || formula.asText().isBlank()) {
            throw ServiceException.validation("drawing:" + payload.path("chartId").asText() + "." + field + " must be a non-empty formula");
        }
        return formula;
    }

    private ObjectNode chartTextFormulaModel(ObjectNode payload, String field) {
        JsonNode rawElements = payload.get("elements");
        if (rawElements == null || !rawElements.isObject()) {
            throw ServiceException.validation("Chart elements must be an object when linked text formulas are inspected");
        }
        ObjectNode elements = (ObjectNode) rawElements;
        String ownerField = switch (field) {
            case "titleText.linkedFormula" -> "titleText";
            case "legend.text.linkedFormula" -> "legend.text";
            case "categoryAxis.titleText.linkedFormula" -> "categoryAxis.titleText";
            case "valueAxis.titleText.linkedFormula" -> "valueAxis.titleText";
            case "secondaryCategoryAxis.titleText.linkedFormula" -> "secondaryCategoryAxis.titleText";
            case "secondaryValueAxis.titleText.linkedFormula" -> "secondaryValueAxis.titleText";
            case "dataTable.font.linkedFormula" -> "dataTable.font";
            default -> throw ServiceException.validation("Unsupported chart text formula field: " + field);
        };
        String[] path = ownerField.split("\\.");
        ObjectNode current = elements;
        for (String segment : path) {
            JsonNode child = current.get(segment);
            if (child == null) return JsonNodeFactory.instance.objectNode();
            if (child.isNull()) throw ServiceException.validation("Chart text formula owner must not be null: " + field);
            if (!child.isObject()) throw ServiceException.validation("Chart text formula owner must be an object: " + field);
            current = (ObjectNode) child;
        }
        return current;
    }

    private String renameFormula(String formula, String previousName, String nextName) {
        return FormulaReferenceTransformer.renameSheet(formula, previousName, nextName);
    }

    private void rewriteFormulaMetadataSource(ObjectNode metadata, String previousName, String nextName, String operation) {
        JsonNode sourceFormula = metadata.get("sourceFormula");
        if (sourceFormula == null || !sourceFormula.isTextual()) return;
        String rewritten = renameFormula(sourceFormula.asText(), previousName, nextName);
        if (!rewritten.equals(sourceFormula.asText()) && metadata.path("preservedOnly").asBoolean(false)) {
            throw ServiceException.unavailable("UNSUPPORTED_STRUCTURAL_REFERENCE: preserved-only formula references " + operation + " worksheet");
        }
        metadata.put("sourceFormula", rewritten);
    }

    private JsonNode optionalArray(JsonNode owner, String field) {
        JsonNode values = owner == null ? null : owner.get(field);
        if (values == null || values.isNull()) return null;
        if (!values.isArray()) throw ServiceException.validation(field + " must be an array");
        return values;
    }
}
