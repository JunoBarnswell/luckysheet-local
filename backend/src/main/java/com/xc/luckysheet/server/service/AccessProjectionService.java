package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.RangeAccessRegion;
import com.xc.luckysheet.server.contract.RangeRef;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Subject-specific read projection for workbook snapshots and revision delivery. */
@Service
public class AccessProjectionService {
    private record HiddenDataReferences(Set<String> dataSourceIds, Set<String> workbookTableIds,
                                       Set<String> sheetTableIds, Set<String> pivotIds,
                                       Set<String> namedRangeNames) { }

    private static final Pattern CELL_REFERENCE = Pattern.compile(
            "(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?\\$?([A-Z]{1,3})\\$?([1-9][0-9]*)(?::\\$?([A-Z]{1,3})\\$?([1-9][0-9]*))?(?![A-Za-z0-9_\\[])(?!\\s*\\()",
            Pattern.CASE_INSENSITIVE);
    private static final Pattern STRING_LITERAL = Pattern.compile("\"(?:[^\"]|\"\")*\"");
    private static final Pattern FUNCTION_NAME = Pattern.compile("(?i)\\b[A-Z_][A-Z0-9_.]*\\s*\\(");
    private static final Pattern ERROR_LITERAL = Pattern.compile("#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\\?|NUM!|N/A|CALC!|BLOCKED!|SPILL!|PARSE!)", Pattern.CASE_INSENSITIVE);
    private static final Set<String> BOOLEAN_LITERALS = Set.of("TRUE", "FALSE");
    private static final Set<String> RANGELESS_METADATA_MUTATIONS = Set.of(
            "sheet.add", "sheet.duplicated", "sheet.hidden", "sheet.rename", "sheet.reordered", "sheet.tabColor", "sheet.unhidden",
            "workbook.calculation.mode.set", "workbook.editing.options.set", "workbook.renamed",
            "view.set", "pageLayout.margins.set", "pageLayout.orientation.set", "pageLayout.paperSize.set",
            "pageLayout.pageSetupDetail.set", "pageLayout.scaleToFit.set", "pageLayout.printTitles.set",
            "pageLayout.printArea.set", "pageLayout.printArea.clear", "pageLayout.pageBreak.insert",
            "pageLayout.pageBreak.remove", "pageLayout.pageBreak.clear", "pageLayout.printGridlines.set",
            "pageLayout.printHeadings.set", "pageLayout.viewGridlines.set", "pageLayout.viewHeadings.set"
    );

    public AccessProjectionService(ObjectMapper mapper) {
        // Keep constructor injection so Spring owns a single projection service.
    }

    public JsonNode projectSnapshot(JsonNode canonical, RangeAccessResolver access) {
        JsonNode projected = canonical.deepCopy();
        List<RangeAccessRegion> hiddenRegions = access.hiddenRegions();
        if (hiddenRegions.isEmpty()) return projected;
        if (!projected.isObject() || !projected.path("sheets").isArray()) {
            throw new ServiceException("STORAGE_CORRUPT", 500, "Canonical workbook snapshot has no sheet collection");
        }

        Map<String, String> sheetNames = new HashMap<>();
        for (JsonNode sheet : projected.path("sheets")) {
            if (sheet.path("id").isTextual() && sheet.path("name").isTextual()) {
                sheetNames.put(sheet.path("name").asText().toLowerCase(java.util.Locale.ROOT), sheet.path("id").asText());
            }
        }
        HiddenDataReferences hiddenReferences = redactDataModel(projected, access, sheetNames);
        for (JsonNode rawSheet : projected.path("sheets")) {
            if (!(rawSheet instanceof ObjectNode sheet)) continue;
            String sheetId = sheet.path("id").asText("");
            JsonNode rawCells = sheet.get("cells");
            if (rawCells instanceof ObjectNode cells) redactCells(sheetId, cells, access, sheetNames);
            redactCellKeyedReview(sheet, sheetId, access);
            redactHyperlinks(sheet, sheetId, access);
            redactRangeBoundDerivedObjects(sheet, access, sheetNames, hiddenReferences);
        }
        return projected;
    }

    /** Full mutation parameters are safe only when every server-resolved range is readable. */
    public boolean canDeliver(CommittedOperationEnvelope operation, RangeAccessResolver access) {
        if (access.hiddenRegions().isEmpty()) return true;
        for (CommittedOperationMutation mutation : operation.mutations()) {
            List<RangeRef> ranges = new ArrayList<>(mutation.affectedRanges());
            ranges.addAll(mutation.structuralImpactRanges());
            if (ranges.isEmpty() && mutation.params() != null && !mutation.params().isNull() && !mutation.params().isEmpty()
                    && !RANGELESS_METADATA_MUTATIONS.contains(mutation.id())) return false;
            for (RangeRef range : ranges) if (!access.canRead(range)) return false;
            // A visible destination does not authorize exposing a formula's
            // cached value when one of its dependencies is hidden.
            if (containsFormulaReferenceToHiddenData(mutation.params(), mutation.sheetId(), access, Map.of())) return false;
        }
        return true;
    }

    /** Formula writes are rejected when a dependency is unreadable by this subject. */
    public boolean formulaDependenciesReadable(JsonNode canonical, JsonNode payload, String currentSheetId,
                                              RangeAccessResolver access) {
        if (access.hiddenRegions().isEmpty()) return true;
        Map<String, String> sheetNames = new HashMap<>();
        for (JsonNode sheet : canonical.path("sheets")) {
            if (sheet.path("id").isTextual() && sheet.path("name").isTextual()) {
                sheetNames.put(sheet.path("name").asText().toLowerCase(java.util.Locale.ROOT), sheet.path("id").asText());
            }
        }
        return !containsFormulaReferenceToHiddenData(payload, currentSheetId, access, sheetNames);
    }

    /** Presence payloads with drafts or cell values are delivered only after checking addressed ranges. */
    public boolean canDeliverEphemeral(JsonNode state, RangeAccessResolver access) {
        if (access.hiddenRegions().isEmpty()) return true;
        if (containsUnaddressedPayload(state)) return false;
        return allAddressesReadable(state, access);
    }

    private static HiddenDataReferences redactDataModel(JsonNode root, RangeAccessResolver access,
                                                         Map<String, String> sheetNames) {
        Set<String> hiddenDataSources = new HashSet<>();
        Set<String> hiddenWorkbookTables = new HashSet<>();
        Set<String> hiddenSheetTables = new HashSet<>();
        Set<String> hiddenPivots = new HashSet<>();
        Set<String> hiddenNamedRanges = new HashSet<>();

        // A block-backed source may omit its own range and rely on the sheet's
        // data-region binding. Resolve that binding before filtering sources.
        for (JsonNode sheet : root.path("sheets")) {
            for (JsonNode region : sheet.path("dataRegions")) {
                if (containsUnreadableRange(region, access)) {
                    hiddenDataSources.add(region.path("sourceId").asText(""));
                }
            }
        }

        JsonNode dataModelNode = root.get("dataModel");
        if (dataModelNode instanceof ObjectNode dataModel) {
            filterArray(dataModel.get("sources"), source -> {
                String id = source.path("id").asText("");
                if (containsUnreadableRange(source, access)) hiddenDataSources.add(id);
                return !hiddenDataSources.contains(id);
            });
            // A table can bind directly to a range or to a block-backed source.
            filterArray(dataModel.get("tables"), table -> {
                String id = table.path("id").asText("");
                boolean hidden = containsUnreadableRange(table, access)
                        || hiddenDataSources.contains(table.path("sourceId").asText(""));
                if (hidden) hiddenWorkbookTables.add(id);
                return !hidden;
            });
            filterArray(dataModel.get("relationships"), relationship ->
                    !hiddenWorkbookTables.contains(relationship.path("fromTableId").asText(""))
                            && !hiddenWorkbookTables.contains(relationship.path("toTableId").asText("")));
            filterArray(dataModel.get("views"), view -> !hiddenWorkbookTables.contains(view.path("tableId").asText("")));
        }

        for (JsonNode sheet : root.path("sheets")) {
            for (JsonNode table : sheet.path("sheetTables")) {
                if (containsUnreadableRange(table, access)) hiddenSheetTables.add(table.path("id").asText(""));
            }
        }
        filterArray(root.get("definedNameModels"), name -> {
            String formula = name.path("formula").asText("");
            String ownerSheet = name.path("sheetId").asText("");
            boolean hidden = !formula.isBlank() && referencesHiddenData(formula, ownerSheet, access, sheetNames);
            if (hidden) hiddenNamedRanges.add(name.path("name").asText("").toLowerCase(java.util.Locale.ROOT));
            return !hidden;
        });
        JsonNode projectedNames = root.get("definedNames");
        if (projectedNames instanceof ObjectNode names) {
            List<String> hiddenProjectionKeys = new ArrayList<>();
            names.fieldNames().forEachRemaining(name -> {
                if (hiddenNamedRanges.contains(name.toLowerCase(java.util.Locale.ROOT))) hiddenProjectionKeys.add(name);
            });
            hiddenProjectionKeys.forEach(names::remove);
        }

        for (JsonNode sheetNode : root.path("sheets")) {
            if (!(sheetNode instanceof ObjectNode sheet) || !(sheet.get("pivots") instanceof ArrayNode pivots)) continue;
            String sheetId = sheet.path("id").asText("");
            List<Integer> removals = new ArrayList<>();
            for (int i = 0; i < pivots.size(); i++) {
                JsonNode pivot = pivots.get(i);
                JsonNode source = pivot.path("source");
                String kind = source.path("kind").asText("");
                boolean hiddenSource = switch (kind) {
                    case "data-source" -> hiddenDataSources.contains(source.path("dataSourceId").asText(""));
                    case "table" -> hiddenWorkbookTables.contains(source.path("tableId").asText(""))
                            || hiddenSheetTables.contains(source.path("tableId").asText(""));
                    case "named-range" -> hiddenNamedRanges.contains(source.path("name").asText("").toLowerCase(java.util.Locale.ROOT));
                    case "worksheet-range", "worksheet-ranges" -> false; // handled by the range scan below
                    default -> true;
                };
                if (hiddenSource || containsUnreadableRange(pivot, access)) {
                    removals.add(i);
                    if (pivot.path("id").isTextual()) hiddenPivots.add(pivot.path("id").asText());
                }
            }
            for (int i = removals.size() - 1; i >= 0; i--) pivots.remove(removals.get(i));
        }

        filterArray(root.get("queryDefinitions"), definition -> {
            JsonNode target = definition.path("lastTarget");
            if (queryTargetIsUnreadable(target, access)) return false;
            if (hiddenWorkbookTables.contains(target.path("tableId").asText(""))
                    || hiddenSheetTables.contains(target.path("tableId").asText(""))
                    || hiddenPivots.contains(target.path("pivotId").asText(""))) return false;
            return true;
        });

        return new HiddenDataReferences(Set.copyOf(hiddenDataSources), Set.copyOf(hiddenWorkbookTables),
                Set.copyOf(hiddenSheetTables), Set.copyOf(hiddenPivots), Set.copyOf(hiddenNamedRanges));
    }

    private static void filterArray(JsonNode node, java.util.function.Predicate<JsonNode> keep) {
        if (!(node instanceof ArrayNode array)) return;
        List<Integer> removals = new ArrayList<>();
        for (int i = 0; i < array.size(); i++) if (!keep.test(array.get(i))) removals.add(i);
        for (int i = removals.size() - 1; i >= 0; i--) array.remove(removals.get(i));
    }

    private static boolean queryTargetIsUnreadable(JsonNode target, RangeAccessResolver access) {
        JsonNode range = target == null ? null : target.get("range");
        JsonNode sheetId = target == null ? null : target.get("sheetId");
        if (range == null || !sheetIdIsTextual(sheetId)
                || !range.path("startRow").canConvertToInt() || !range.path("startColumn").canConvertToInt()) return false;
        int startRow = range.path("startRow").asInt();
        int startColumn = range.path("startColumn").asInt();
        int endRow = range.path("endRow").canConvertToInt() ? range.path("endRow").asInt() : startRow;
        int endColumn = range.path("endColumn").canConvertToInt() ? range.path("endColumn").asInt() : startColumn;
        return !access.canRead(new RangeRef(sheetId.asText(), startRow, endRow, startColumn, endColumn));
    }

    private static boolean sheetIdIsTextual(JsonNode value) {
        return value != null && value.isTextual() && !value.asText().isBlank();
    }

    private static boolean containsUnaddressedPayload(JsonNode node) {
        if (node == null || node.isNull()) return false;
        if (node.isObject()) {
            if (node.has("draft") || node.has("formula") || node.has("value") || node.has("cellText") || node.has("content")) return true;
            Iterator<JsonNode> children = node.elements();
            while (children.hasNext()) if (containsUnaddressedPayload(children.next())) return true;
        } else if (node.isArray()) {
            for (JsonNode child : node) if (containsUnaddressedPayload(child)) return true;
        }
        return false;
    }

    private static boolean allAddressesReadable(JsonNode node, RangeAccessResolver access) {
        if (node == null || node.isNull()) return true;
        if (node.isObject()) {
            JsonNode sheetId = node.get("sheetId");
            JsonNode row = node.has("row") ? node.get("row") : node.get("startRow");
            JsonNode column = node.has("column") ? node.get("column") : node.get("startColumn");
            if (sheetId != null && sheetId.isTextual() && row != null && row.canConvertToInt()
                    && column != null && column.canConvertToInt()) {
                JsonNode endRow = node.has("endRow") ? node.get("endRow") : row;
                JsonNode endColumn = node.has("endColumn") ? node.get("endColumn") : column;
                if (endRow == null || !endRow.canConvertToInt() || endColumn == null || !endColumn.canConvertToInt()) return false;
                if (!access.canRead(new RangeRef(sheetId.asText(), row.asInt(), endRow.asInt(), column.asInt(), endColumn.asInt()))) return false;
            }
            Iterator<JsonNode> children = node.elements();
            while (children.hasNext()) if (!allAddressesReadable(children.next(), access)) return false;
        } else if (node.isArray()) {
            for (JsonNode child : node) if (!allAddressesReadable(child, access)) return false;
        }
        return true;
    }

    private static void redactCells(String currentSheetId, ObjectNode cells, RangeAccessResolver access,
                                    Map<String, String> sheetNames) {
        List<String> emptyRows = new ArrayList<>();
        Iterator<Map.Entry<String, JsonNode>> rows = cells.fields();
        while (rows.hasNext()) {
            Map.Entry<String, JsonNode> row = rows.next();
            int rowIndex = parseCoordinate(row.getKey());
            if (rowIndex < 0 || !(row.getValue() instanceof ObjectNode columns)) continue;
            List<String> removeColumns = new ArrayList<>();
            Iterator<Map.Entry<String, JsonNode>> columnEntries = columns.fields();
            while (columnEntries.hasNext()) {
                Map.Entry<String, JsonNode> column = columnEntries.next();
                int columnIndex = parseCoordinate(column.getKey());
                if (columnIndex < 0) continue;
                if (!access.canRead(new RangeRef(currentSheetId, rowIndex, rowIndex, columnIndex, columnIndex))) {
                    removeColumns.add(column.getKey());
                } else if (column.getValue() instanceof ObjectNode cell
                        && containsFormulaReferenceToHiddenData(cell, currentSheetId, access, sheetNames)) {
                    redactFormula(cell);
                }
            }
            removeColumns.forEach(columns::remove);
            if (columns.isEmpty()) emptyRows.add(row.getKey());
        }
        emptyRows.forEach(cells::remove);
    }

    private static boolean referencesHiddenData(String formula, String currentSheetId, RangeAccessResolver access,
                                                Map<String, String> sheetNames) {
        String sanitized = STRING_LITERAL.matcher(formula).replaceAll("");
        if (Pattern.compile("(?i)\\b(?:INDIRECT|OFFSET)\\s*\\(").matcher(sanitized).find() || sanitized.indexOf('[') >= 0) return true;
        Matcher references = CELL_REFERENCE.matcher(sanitized);
        StringBuffer withoutReferences = new StringBuffer();
        while (references.find()) {
            String sheetToken = references.group(1) != null ? references.group(1).replace("''", "'") : references.group(2);
            String referencedSheet = sheetToken == null ? currentSheetId : sheetNames.get(sheetToken.trim().toLowerCase(java.util.Locale.ROOT));
            if (referencedSheet == null) return true;
            int startColumn = columnIndex(references.group(3));
            int startRow = Integer.parseInt(references.group(4)) - 1;
            int endColumn = references.group(5) == null ? startColumn : columnIndex(references.group(5));
            int endRow = references.group(6) == null ? startRow : Integer.parseInt(references.group(6)) - 1;
            if (!access.canRead(new RangeRef(referencedSheet, startRow, endRow, startColumn, endColumn))) return true;
            references.appendReplacement(withoutReferences, "");
        }
        references.appendTail(withoutReferences);
        String residue = FUNCTION_NAME.matcher(withoutReferences).replaceAll("");
        residue = ERROR_LITERAL.matcher(residue).replaceAll("");
        for (String literal : BOOLEAN_LITERALS) residue = residue.replaceAll("(?i)\\b" + literal + "\\b", "");
        residue = residue.replaceAll("\\d+(?:\\.\\d+)?", "").replaceAll("[\\s=+*/^&%(),;:{}!<>?~.-]", "");
        return residue.chars().anyMatch(Character::isLetter);
    }

    private static void redactFormula(ObjectNode cell) {
        cell.put("formula", "=#BLOCKED!");
        cell.put("value", (String) null);
        cell.remove("displayValue");
        cell.remove("richText");
        cell.remove("formulaMetadata");
        // Presentation payloads can preserve formula source text and cached render values.
        cell.remove("presentation");
        cell.set("formulaValue", mapperNode("#BLOCKED!"));
    }

    private static boolean containsFormulaReferenceToHiddenData(JsonNode node, String currentSheetId,
                                                                 RangeAccessResolver access,
                                                                 Map<String, String> sheetNames) {
        if (node == null || node.isNull()) return false;
        if (node.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> fields = node.fields();
            while (fields.hasNext()) {
                Map.Entry<String, JsonNode> field = fields.next();
                JsonNode value = field.getValue();
                String key = field.getKey();
                boolean formulaField = key.equals("formula") || key.equals("sourceFormula")
                        || key.equals("formula1") || key.equals("formula2");
                boolean formulaValuedRuleField = (key.equals("value1") || key.equals("value2"))
                        && value.isTextual() && value.asText().startsWith("=");
                if ((formulaField || formulaValuedRuleField) && value.isTextual()
                        && referencesHiddenData(value.asText(), currentSheetId, access, sheetNames)) return true;
                if (containsFormulaReferenceToHiddenData(value, currentSheetId, access, sheetNames)) return true;
            }
        } else if (node.isArray()) {
            for (JsonNode child : node) {
                if (containsFormulaReferenceToHiddenData(child, currentSheetId, access, sheetNames)) return true;
            }
        }
        return false;
    }

    private static JsonNode mapperNode(String code) {
        return com.fasterxml.jackson.databind.node.JsonNodeFactory.instance.objectNode()
                .put("kind", "error").put("code", code).put("message", "Formula references data the current subject cannot view");
    }

    private static void redactCellKeyedReview(ObjectNode sheet, String sheetId, RangeAccessResolver access) {
        JsonNode reviewNode = sheet.get("review");
        if (!(reviewNode instanceof ObjectNode review)) return;
        Set<String> removedThreadIds = new HashSet<>();
        removeHiddenCellKeys(review.get("notesByCell"), sheetId, access, null);
        removeHiddenCellKeys(review.get("threadIdsByCell"), sheetId, access, removedThreadIds);
        JsonNode threadMap = review.get("threadsById");
        if (threadMap instanceof ObjectNode threads) removedThreadIds.forEach(threads::remove);
    }

    private static void removeHiddenCellKeys(JsonNode node, String sheetId, RangeAccessResolver access, Set<String> removedIds) {
        if (!(node instanceof ObjectNode entries)) return;
        List<String> toRemove = new ArrayList<>();
        Iterator<Map.Entry<String, JsonNode>> fields = entries.fields();
        while (fields.hasNext()) {
            Map.Entry<String, JsonNode> entry = fields.next();
            int[] cell = parseCellKey(entry.getKey());
            if (cell != null && !access.canRead(new RangeRef(sheetId, cell[0], cell[0], cell[1], cell[1]))) {
                toRemove.add(entry.getKey());
                if (removedIds != null && entry.getValue().isArray()) {
                    entry.getValue().forEach(value -> { if (value.isTextual()) removedIds.add(value.asText()); });
                }
            }
        }
        toRemove.forEach(entries::remove);
    }

    private static int[] parseCellKey(String key) {
        Matcher matcher = Pattern.compile("(?:^|[!|])([0-9]+)[:,]([0-9]+)$").matcher(key);
        if (!matcher.find()) return null;
        try { return new int[] { Integer.parseInt(matcher.group(1)), Integer.parseInt(matcher.group(2)) }; }
        catch (NumberFormatException ignored) { return null; }
    }

    private static void redactHyperlinks(ObjectNode sheet, String sheetId, RangeAccessResolver access) {
        JsonNode raw = sheet.get("hyperlinks");
        if (!(raw instanceof ArrayNode hyperlinks)) return;
        List<Integer> removals = new ArrayList<>();
        for (int i = 0; i < hyperlinks.size(); i++) {
            JsonNode link = hyperlinks.get(i);
            if (!link.path("row").canConvertToInt() || !link.path("column").canConvertToInt()) continue;
            int row = link.path("row").asInt();
            int column = link.path("column").asInt();
            if (!access.canRead(new RangeRef(sheetId, row, row, column, column))) removals.add(i);
        }
        for (int i = removals.size() - 1; i >= 0; i--) hyperlinks.remove(removals.get(i));
    }

    private static void redactRangeBoundDerivedObjects(ObjectNode sheet, RangeAccessResolver access,
                                                       Map<String, String> sheetNames,
                                                       HiddenDataReferences hiddenReferences) {
        for (String property : List.of("pivots", "sparklines", "sheetTables", "tables", "dataRegions")) {
            JsonNode raw = sheet.get(property);
            if (!(raw instanceof ArrayNode items)) continue;
            List<Integer> removals = new ArrayList<>();
            for (int i = 0; i < items.size(); i++) if (containsUnreadableRange(items.get(i), access)) removals.add(i);
            for (int i = removals.size() - 1; i >= 0; i--) items.remove(removals.get(i));
        }

        String currentSheetId = sheet.path("id").asText("");
        for (String property : List.of("conditionalFormats", "dataValidations")) {
            JsonNode raw = sheet.get(property);
            if (!(raw instanceof ArrayNode rules)) continue;
            List<Integer> removals = new ArrayList<>();
            for (int i = 0; i < rules.size(); i++) {
                JsonNode rule = rules.get(i);
                if (!hasFullyReadableRuleRanges(rule, access) || containsUnreadableRange(rule, access)
                        || containsFormulaReferenceToHiddenData(rule, currentSheetId, access, sheetNames)) {
                    removals.add(i);
                }
            }
            for (int i = removals.size() - 1; i >= 0; i--) rules.remove(removals.get(i));
        }

        JsonNode rawDrawings = sheet.get("drawings");
        JsonNode rawPayloads = sheet.get("drawingPayloads");
        if (!(rawDrawings instanceof ArrayNode drawings) || !(rawPayloads instanceof ObjectNode payloads)) return;
        Set<String> referencedPayloadIds = new HashSet<>();
        for (int i = drawings.size() - 1; i >= 0; i--) {
            JsonNode drawing = drawings.get(i);
            String payloadId = drawing.path("payloadId").asText("");
            JsonNode payload = payloads.get(payloadId);
            if (drawingIntersectsUnreadableRange(drawing, access) || containsUnreadableRange(payload, access)
                    || referencesHiddenDataModel(drawing, hiddenReferences)
                    || referencesHiddenDataModel(payload, hiddenReferences)) {
                drawings.remove(i);
                continue;
            }
            if (!payloadId.isBlank()) referencedPayloadIds.add(payloadId);
        }
        List<String> unusedPayloadIds = new ArrayList<>();
        payloads.fieldNames().forEachRemaining(payloadId -> {
            if (!referencedPayloadIds.contains(payloadId)) unusedPayloadIds.add(payloadId);
        });
        unusedPayloadIds.forEach(payloads::remove);
    }

    private static boolean referencesHiddenDataModel(JsonNode node, HiddenDataReferences hiddenReferences) {
        if (node == null || node.isNull()) return false;
        if (node.isObject()) {
            Iterator<Map.Entry<String, JsonNode>> fields = node.fields();
            while (fields.hasNext()) {
                Map.Entry<String, JsonNode> field = fields.next();
                String key = field.getKey();
                JsonNode value = field.getValue();
                if (value.isTextual()) {
                    String id = value.asText();
                    if ((key.equals("pivotId") && hiddenReferences.pivotIds().contains(id))
                            || ((key.equals("tableId") || key.equals("table"))
                                && (hiddenReferences.workbookTableIds().contains(id) || hiddenReferences.sheetTableIds().contains(id)))
                            || ((key.equals("dataSourceId") || key.equals("sourceId"))
                                && hiddenReferences.dataSourceIds().contains(id))) return true;
                }
                if (referencesHiddenDataModel(value, hiddenReferences)) return true;
            }
        } else if (node.isArray()) {
            for (JsonNode child : node) if (referencesHiddenDataModel(child, hiddenReferences)) return true;
        }
        return false;
    }

    private static boolean hasFullyReadableRuleRanges(JsonNode rule, RangeAccessResolver access) {
        JsonNode ranges = rule == null ? null : rule.get("ranges");
        if (!(ranges instanceof ArrayNode rangeArray) || rangeArray.isEmpty()) return false;
        for (JsonNode range : rangeArray) {
            if (!range.path("sheetId").isTextual() || !range.path("startRow").canConvertToInt()
                    || !range.path("endRow").canConvertToInt() || !range.path("startColumn").canConvertToInt()
                    || !range.path("endColumn").canConvertToInt()) return false;
            if (!access.canRead(new RangeRef(range.path("sheetId").asText(), range.path("startRow").asInt(),
                    range.path("endRow").asInt(), range.path("startColumn").asInt(), range.path("endColumn").asInt()))) {
                return false;
            }
        }
        return true;
    }

    private static boolean drawingIntersectsUnreadableRange(JsonNode drawing, RangeAccessResolver access) {
        JsonNode anchor = drawing == null ? null : drawing.get("anchor");
        if (anchor == null || !anchor.isObject() || !anchor.path("row").canConvertToInt()
                || !anchor.path("column").canConvertToInt()) return false;
        int startRow = anchor.path("row").asInt();
        int startColumn = anchor.path("column").asInt();
        int endRow = anchor.path("endRow").canConvertToInt() ? anchor.path("endRow").asInt() : startRow;
        int endColumn = anchor.path("endColumn").canConvertToInt() ? anchor.path("endColumn").asInt() : startColumn;
        String sheetId = drawing.path("sheetId").asText("");
        if (sheetId.isBlank()) return true;
        return !access.canRead(new RangeRef(sheetId, Math.min(startRow, endRow), Math.max(startRow, endRow),
                Math.min(startColumn, endColumn), Math.max(startColumn, endColumn)));
    }

    private static boolean containsUnreadableRange(JsonNode node, RangeAccessResolver access) {
        if (node == null || node.isNull()) return false;
        if (node.isObject()) {
            JsonNode sheet = node.get("sheetId");
            JsonNode startRow = node.get("startRow");
            JsonNode endRow = node.get("endRow");
            JsonNode startColumn = node.get("startColumn");
            JsonNode endColumn = node.get("endColumn");
            if (sheet != null && sheet.isTextual() && startRow != null && startRow.canConvertToInt()
                    && endRow != null && endRow.canConvertToInt() && startColumn != null && startColumn.canConvertToInt()
                    && endColumn != null && endColumn.canConvertToInt()) {
                RangeRef range = new RangeRef(sheet.asText(), startRow.asInt(), endRow.asInt(), startColumn.asInt(), endColumn.asInt());
                if (!access.canRead(range)) return true;
            }
            Iterator<JsonNode> children = node.elements();
            while (children.hasNext()) if (containsUnreadableRange(children.next(), access)) return true;
        } else if (node.isArray()) {
            for (JsonNode child : node) if (containsUnreadableRange(child, access)) return true;
        }
        return false;
    }

    private static int parseCoordinate(String value) {
        try {
            int coordinate = Integer.parseInt(value);
            return coordinate < 0 ? -1 : coordinate;
        } catch (NumberFormatException ignored) {
            return -1;
        }
    }

    private static int columnIndex(String letters) {
        int column = 0;
        for (int i = 0; i < letters.length(); i++) column = column * 26 + Character.toUpperCase(letters.charAt(i)) - 'A' + 1;
        return column - 1;
    }
}
