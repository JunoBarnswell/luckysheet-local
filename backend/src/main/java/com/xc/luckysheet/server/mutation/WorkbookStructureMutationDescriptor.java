package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
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
    static final Set<String> IDS = Set.of(
            "sheet.add", "sheet.remove", "sheet.rename", "sheet.duplicated", "sheet.restore",
            "hyperlink.set", "hyperlink.remove"
    );

    WorkbookStructureMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, id.startsWith("hyperlink."), id.startsWith("hyperlink.") ? "edit-cell" : "format");
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
        validateNoExternalSheetReferences(root, id, sheets.get(index).path("name").asText());
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

    private void cloneSheetScopedDocuments(ObjectNode root, String sourceSheetId, String sourceName, String targetSheetId, String targetName) {
        JsonNode names = root.get("definedNameModels");
        if (names != null && names.isArray()) {
            List<JsonNode> copies = new ArrayList<>();
            for (JsonNode raw : names) {
                if (!raw.isObject() || !"sheet".equals(raw.path("scope").asText()) || !sourceSheetId.equals(raw.path("sheetId").asText())) continue;
                ObjectNode copy = ((ObjectNode) raw).deepCopy();
                copy.put("sheetId", targetSheetId);
                if (copy.path("formula").isTextual()) copy.put("formula", FormulaReferenceTransformer.renameSheet(copy.path("formula").asText(), sourceName, targetName));
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
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            ObjectNode sheet = (ObjectNode) rawSheet;
            if (sourceSheetId.equals(sheet.path("id").asText())) continue;
            JsonNode cells = sheet.get("cells");
            if (cells != null && cells.isObject()) cells.fields().forEachRemaining(row -> {
                if (!row.getValue().isObject()) return;
                row.getValue().fields().forEachRemaining(cell -> {
                    JsonNode formula = cell.getValue().get("formula");
                    if (formula != null && formula.isTextual() && !formula.asText().equals(FormulaReferenceTransformer.renameSheet(formula.asText(), sourceName, sourceName + "__deleted__"))) references.add("cell-formula:" + formula.asText());
                });
            });
            JsonNode hyperlinks = sheet.get("hyperlinks");
            if (hyperlinks != null && hyperlinks.isArray()) for (JsonNode hyperlink : hyperlinks) if (sourceSheetId.equals(hyperlink.path("hyperlink").path("target").path("sheetId").asText())) references.add("hyperlink:" + hyperlink.path("hyperlink").path("id").asText());
            for (String field : List.of("merges", "conditionalFormats", "dataValidations", "dataRegions", "sheetTables", "spillRanges", "protectionRules")) {
                JsonNode values = sheet.get(field);
                if (values != null && values.isArray()) for (JsonNode value : values) if (containsDeletedRange(value, sourceSheetId)) references.add(field + ":" + value.path("id").asText());
            }
            JsonNode pivots = sheet.get("pivots");
            if (pivots != null && pivots.isArray()) for (JsonNode pivot : pivots) if (containsDeletedPivotReference(pivot, sourceSheetId)) references.add("pivot:" + pivot.path("id").asText());
            JsonNode sparklines = sheet.get("sparklines");
            if (sparklines != null && sparklines.isArray()) for (JsonNode sparkline : sparklines) if (containsDeletedRange(sparkline, sourceSheetId)) references.add("sparkline:" + sparkline.path("id").asText());
        }
        JsonNode names = root.get("definedNameModels");
        if (names != null && names.isArray()) for (JsonNode name : names) if (!("sheet".equals(name.path("scope").asText()) && sourceSheetId.equals(name.path("sheetId").asText())) && name.path("formula").isTextual() && !name.path("formula").asText().equals(FormulaReferenceTransformer.renameSheet(name.path("formula").asText(), sourceName, sourceName + "__deleted__"))) references.add("defined-name:" + name.path("name").asText());
        if (!references.isEmpty()) throw ServiceException.conflict("Cannot delete sheet with external references: " + String.join(", ", references));
    }

    private boolean containsDeletedRange(JsonNode value, String sourceSheetId) {
        if (!value.isObject()) return false;
        if (sourceSheetId.equals(value.path("sheetId").asText())) return true;
        if (sourceSheetId.equals(value.path("range").path("sheetId").asText())) return true;
        if (sourceSheetId.equals(value.path("sourceRange").path("sheetId").asText())) return true;
        if (sourceSheetId.equals(value.path("formulaAnchor").path("sheetId").asText())) return true;
        JsonNode ranges = value.get("ranges");
        if (ranges != null && ranges.isArray()) for (JsonNode range : ranges) if (sourceSheetId.equals(range.path("sheetId").asText())) return true;
        JsonNode listSource = value.path("listSource");
        return sourceSheetId.equals(listSource.path("range").path("sheetId").asText());
    }

    private boolean containsDeletedPivotReference(JsonNode value, String sourceSheetId) {
        if (!value.isObject()) return false;
        if (sourceSheetId.equals(value.path("target").path("sheetId").asText())) return true;
        JsonNode source = value.get("source");
        if (source == null || !source.isObject()) return false;
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
        remapSheetOwnedObjects(copy, sourceSheetId, targetSheetId);
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
        ObjectNode remapped = values.objectNode();
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
        if ("camera".equals(kind)) remapRange(payload.get("sourceRange"), sourceSheetId, targetSheetId);
        if ("chart".equals(kind)) {
            remapRange(payload.get("categoryRange"), sourceSheetId, targetSheetId);
            JsonNode ranges = payload.get("sourceRanges");
            if (ranges != null && ranges.isArray()) for (JsonNode range : ranges) remapRange(range, sourceSheetId, targetSheetId);
            if (payload.path("pivotId").isTextual()) payload.put("pivotId", pivotIds.getOrDefault(payload.path("pivotId").asText(), payload.path("pivotId").asText()));
        }
        if ("data-chart".equals(kind) && payload.path("source").isObject()) {
            ObjectNode source = (ObjectNode) payload.get("source");
            remapRange(source.get("range"), sourceSheetId, targetSheetId);
            if (source.path("tableId").isTextual()) source.put("tableId", tableIds.getOrDefault(source.path("tableId").asText(), source.path("tableId").asText()));
        }
        if ("connector".equals(kind)) {
            if (payload.path("start").path("drawingId").isTextual()) ((ObjectNode) payload.get("start")).put("drawingId", drawingIds.getOrDefault(payload.path("start").path("drawingId").asText(), payload.path("start").path("drawingId").asText()));
            if (payload.path("end").path("drawingId").isTextual()) ((ObjectNode) payload.get("end")).put("drawingId", drawingIds.getOrDefault(payload.path("end").path("drawingId").asText(), payload.path("end").path("drawingId").asText()));
        }
        if ("form-control".equals(kind)) {
            remapRange(payload.get("cellLink"), sourceSheetId, targetSheetId);
            remapRange(payload.get("inputRange"), sourceSheetId, targetSheetId);
        }
        if ("slicer".equals(kind) || "timeline".equals(kind)) if (payload.path("pivotId").isTextual()) payload.put("pivotId", pivotIds.getOrDefault(payload.path("pivotId").asText(), payload.path("pivotId").asText()));
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
            if (members != null && members.isArray()) for (int index = 0; index < members.size(); index++) if (members.get(index).isTextual()) members.set(index, JsonNodeFactory.instance.textNode(drawingIds.getOrDefault(members.get(index).asText(), members.get(index).asText())));
        }
    }

    private void remapHyperlinks(ObjectNode copy, String sourceSheetId, String targetSheetId) {
        JsonNode values = copy.get("hyperlinks");
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (raw.isObject() && raw.path("hyperlink").path("target").path("sheetId").asText().equals(sourceSheetId)) ((ObjectNode) raw.path("hyperlink").path("target")).put("sheetId", targetSheetId);
        }
    }

    private void remapSheetOwnedObjects(ObjectNode copy, String sourceSheetId, String targetSheetId) {
        for (String field : List.of("commentThreads", "protectionRules")) {
            JsonNode values = copy.get(field);
            if (values != null && values.isArray()) for (JsonNode raw : values) if (raw.isObject() && raw.has("sheetId")) ((ObjectNode) raw).put("sheetId", targetSheetId);
        }
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
                if (cell.getValue().isObject() && cell.getValue().path("formula").isTextual()) ((ObjectNode) cell.getValue()).put("formula", FormulaReferenceTransformer.renameSheet(cell.getValue().path("formula").asText(), sourceName, targetName));
            });
        });
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
            rewriteRuleFormulaFields(sheet.get("conditionalFormats"), previousName, nextName);
            rewriteRuleFormulaFields(sheet.get("dataValidations"), previousName, nextName);
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

    private void rewriteRuleFormulaFields(JsonNode values, String previousName, String nextName) {
        if (values == null || !values.isArray()) return;
        for (JsonNode raw : values) {
            if (!raw.isObject()) continue;
            ObjectNode rule = (ObjectNode) raw;
            for (String field : List.of("value1", "value2", "formula1", "formula2")) {
                if (rule.path(field).isTextual()) rule.put(field, FormulaReferenceTransformer.renameSheet(rule.path(field).asText(), previousName, nextName));
            }
            JsonNode listSource = rule.get("listSource");
            if (listSource != null && listSource.isObject() && listSource.path("kind").asText().equals("formula") && listSource.path("formula").isTextual()) ((ObjectNode) listSource).put("formula", FormulaReferenceTransformer.renameSheet(listSource.path("formula").asText(), previousName, nextName));
        }
    }
}
