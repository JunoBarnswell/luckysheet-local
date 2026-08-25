package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

/** Reducers for sheet-owned metadata with an explicit JSON snapshot truth. */
final class SheetDataMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "sheet.reordered",
            "row.hidden", "row.unhidden", "rows.unhidden.all", "rows.hidden.restore",
            "column.hidden", "column.unhidden", "columns.unhidden.all", "columns.hidden.restore",
            "autoFilter.set", "autoFilter.remove",
            "cf.add", "cf.remove", "cf.clear",
            "dv.add", "dv.remove", "banded.set", "outline.set",
            "sheetTable.add", "sheetTable.remove", "sheetTable.update", "sheetTable.autoFilter.set",
            "tableSheet.update", "ganttSheet.update"
    );

    SheetDataMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, checksProtection(id), protectionAction(id));
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported sheet metadata mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        return switch (id()) {
            case "sheet.reordered" -> List.of();
            case "row.hidden", "row.unhidden" -> List.of(SnapshotMutationSupport.rowRange(root, mutation.sheetId(), renameIndex(params, "index", "row")));
            case "column.hidden", "column.unhidden" -> List.of(SnapshotMutationSupport.columnRange(root, mutation.sheetId(), renameIndex(params, "index", "column")));
            case "rows.unhidden.all", "rows.hidden.restore", "columns.unhidden.all", "columns.hidden.restore" -> List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
            case "autoFilter.set" -> List.of(filterRange(root, mutation.sheetId(), params));
            case "autoFilter.remove" -> existingFilterRange(root, mutation.sheetId());
            case "cf.add" -> ruleRanges(root, mutation.sheetId(), SnapshotMutationSupport.requiredObject(params, "rule"));
            case "cf.remove" -> existingRuleRanges(root, mutation.sheetId(), "conditionalFormats", SnapshotMutationSupport.text(params, "ruleId"));
            case "cf.clear" -> allRuleRanges(root, mutation.sheetId(), "conditionalFormats");
            case "dv.add" -> ruleRanges(root, mutation.sheetId(), SnapshotMutationSupport.requiredObject(params, "rule"));
            case "dv.remove" -> existingRuleRanges(root, mutation.sheetId(), "dataValidations", SnapshotMutationSupport.text(params, "ruleId"));
            case "banded.set" -> bandedRange(root, mutation.sheetId(), params);
            case "outline.set" -> List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
            case "sheetTable.add", "sheetTable.update" -> List.of(tableRange(root, mutation.sheetId(), params));
            case "tableSheet.update", "ganttSheet.update" -> List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
            case "sheetTable.remove", "sheetTable.autoFilter.set" -> List.of(existingTableRange(root, mutation.sheetId(), SnapshotMutationSupport.text(params, "tableId")));
            default -> throw ServiceException.validation("Unsupported sheet metadata mutation: " + id());
        };
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        switch (id()) {
            case "sheet.reordered" -> reorderSheet(root, mutation, params);
            default -> applyToSheet(root, mutation, params);
        }
        return root;
    }

    private void applyToSheet(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        switch (id()) {
            case "row.hidden" -> setHidden(root, sheet, mutation.sheetId(), params, "hiddenRows", "index", true, true);
            case "row.unhidden" -> setHidden(root, sheet, mutation.sheetId(), params, "hiddenRows", "index", false, true);
            case "rows.unhidden.all" -> sheet.set("hiddenRows", JsonNodeFactory.instance.arrayNode());
            case "rows.hidden.restore" -> restoreHidden(root, sheet, mutation.sheetId(), params, "hiddenRows", true);
            case "column.hidden" -> setHidden(root, sheet, mutation.sheetId(), params, "hiddenColumns", "index", true, false);
            case "column.unhidden" -> setHidden(root, sheet, mutation.sheetId(), params, "hiddenColumns", "index", false, false);
            case "columns.unhidden.all" -> sheet.set("hiddenColumns", JsonNodeFactory.instance.arrayNode());
            case "columns.hidden.restore" -> restoreHidden(root, sheet, mutation.sheetId(), params, "hiddenColumns", false);
            case "autoFilter.set" -> sheet.set("autoFilter", filter(root, mutation.sheetId(), params).deepCopy());
            case "autoFilter.remove" -> sheet.remove("autoFilter");
            case "cf.add" -> upsertRule(root, sheet, mutation.sheetId(), params, "conditionalFormats");
            case "cf.remove" -> removeRule(sheet, params, "conditionalFormats");
            case "cf.clear" -> sheet.set("conditionalFormats", JsonNodeFactory.instance.arrayNode());
            case "dv.add" -> upsertRule(root, sheet, mutation.sheetId(), params, "dataValidations");
            case "dv.remove" -> removeRule(sheet, params, "dataValidations");
            case "banded.set" -> setBanded(root, sheet, mutation.sheetId(), params);
            case "outline.set" -> setOutline(root, sheet, mutation.sheetId(), params);
            case "sheetTable.add", "sheetTable.update" -> upsertSheetTable(root, sheet, mutation.sheetId(), params);
            case "tableSheet.update" -> updateTableSheet(root, sheet, mutation.sheetId(), params);
            case "ganttSheet.update" -> updateGanttSheet(root, sheet, mutation.sheetId(), params);
            case "sheetTable.remove" -> removeSheetTable(sheet, params);
            case "sheetTable.autoFilter.set" -> setTableAutoFilter(root, sheet, mutation.sheetId(), params);
            default -> throw ServiceException.validation("Unsupported sheet metadata mutation: " + id());
        }
    }

    private void reorderSheet(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        ArrayNode sheets = SnapshotMutationSupport.sheets(root);
        int currentIndex = sheetIndex(sheets, mutation.sheetId());
        if (currentIndex < 0) throw ServiceException.notFound("Sheet not found: " + mutation.sheetId());
        JsonNode value = params.get("toIndex");
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt()) throw ServiceException.validation("toIndex must be an integer");
        int targetIndex = Math.max(0, Math.min(value.intValue(), sheets.size() - 1));
        JsonNode sheet = sheets.remove(currentIndex);
        sheets.insert(targetIndex, sheet);
    }

    private void setHidden(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params, String property, String indexProperty, boolean hidden, boolean row) {
        ObjectNode coordinateParams = renameIndex(params, indexProperty, row ? "row" : "column");
        int index = SnapshotMutationSupport.index(root, sheetId, coordinateParams, row ? "row" : "column");
        ArrayNode indices = SnapshotMutationSupport.array(sheet, property);
        ArrayNode next = JsonNodeFactory.instance.arrayNode();
        boolean exists = false;
        for (JsonNode item : indices) {
            if (!item.isIntegralNumber() || item.intValue() < 0) throw ServiceException.validation(property + " contains an invalid index");
            if (item.intValue() == index) {
                exists = true;
                if (hidden) next.add(index);
            } else next.add(item.intValue());
        }
        if (hidden && !exists) next.add(index);
        sheet.set(property, next);
    }

    private void restoreHidden(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params, String property, boolean row) {
        JsonNode input = params.get("indices");
        if (input == null || !input.isArray()) throw ServiceException.validation("indices must be an array");
        ArrayNode next = JsonNodeFactory.instance.arrayNode();
        for (JsonNode item : input) {
            ObjectNode coordinate = JsonNodeFactory.instance.objectNode();
            coordinate.set(row ? "row" : "column", item);
            int index = SnapshotMutationSupport.index(root, sheetId, coordinate, row ? "row" : "column");
            if (!containsIndex(next, index)) next.add(index);
        }
        sheet.set(property, next);
    }

    private ObjectNode filter(ObjectNode root, String sheetId, ObjectNode params) {
        ObjectNode filter = SnapshotMutationSupport.requiredObject(params, "autoFilter");
        SnapshotMutationSupport.requireEntitySheet(filter, sheetId);
        RangeRef range = SnapshotMutationSupport.range(root, filter.get("range"));
        SnapshotMutationSupport.requireSheet(range, sheetId);
        JsonNode columns = filter.get("columns");
        if (columns == null || !columns.isObject()) throw ServiceException.validation("AutoFilter columns must be an object");
        validateFilterColumns(columns, range);
        return filter;
    }

    private void setTableAutoFilter(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        String tableId = SnapshotMutationSupport.text(params, "tableId");
        ObjectNode table = SnapshotMutationSupport.findById(SnapshotMutationSupport.array(sheet, "sheetTables"), tableId);
        if (table == null) throw ServiceException.notFound("Sheet Table not found: " + tableId);
        JsonNode value = params.get("autoFilter");
        if (value == null || value.isNull()) {
            table.remove("autoFilter");
            return;
        }
        if (!value.isObject()) throw ServiceException.validation("Table AutoFilter must be an object");
        ObjectNode filter = (ObjectNode) value;
        RangeRef filterRange = SnapshotMutationSupport.range(root, filter.get("range"));
        RangeRef tableRange = SnapshotMutationSupport.range(root, table.get("range"));
        if (!sameRange(filterRange, tableRange)) throw ServiceException.validation("Table AutoFilter range must equal the Table range");
        filter(root, sheetId, params);
        JsonNode worksheetFilter = sheet.get("autoFilter");
        if (worksheetFilter != null && worksheetFilter.isObject()
                && rangesOverlap(filterRange, SnapshotMutationSupport.range(root, worksheetFilter.get("range")))) {
            throw ServiceException.validation("Table AutoFilter cannot overlap a Worksheet AutoFilter");
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "sheetTables")) {
            if (!raw.isObject() || tableId.equals(raw.path("id").asText())) continue;
            JsonNode other = raw.get("autoFilter");
            if (other != null && other.isObject() && rangesOverlap(filterRange, SnapshotMutationSupport.range(root, other.get("range")))) {
                throw ServiceException.validation("Table AutoFilter ranges cannot overlap");
            }
        }
        table.set("autoFilter", filter.deepCopy());
    }

    private void validateFilterColumns(JsonNode columns, RangeRef range) {
        columns.fields().forEachRemaining(entry -> {
            int key;
            try {
                key = Integer.parseInt(entry.getKey());
            } catch (NumberFormatException error) {
                throw ServiceException.validation("AutoFilter column key is invalid");
            }
            if (!entry.getValue().isObject()) throw ServiceException.validation("AutoFilter column is invalid");
            JsonNode column = entry.getValue();
            if (!column.path("column").canConvertToInt() || column.path("column").intValue() != key
                    || key < range.startColumn() || key > range.endColumn()
                    || !column.path("showButton").isBoolean() || !column.path("hiddenButton").isBoolean()) {
                throw ServiceException.validation("AutoFilter column identity is invalid");
            }
            JsonNode criterion = column.get("criterion");
            if (criterion == null || criterion.isNull()) return;
            if (!criterion.isObject() || !Set.of("values", "custom", "dynamic", "top10", "color", "icon").contains(criterion.path("kind").asText())) {
                throw ServiceException.validation("AutoFilter criterion kind is invalid");
            }
        });
    }

    private static boolean sameRange(RangeRef left, RangeRef right) {
        return left.sheetId().equals(right.sheetId()) && left.startRow() == right.startRow() && left.endRow() == right.endRow()
                && left.startColumn() == right.startColumn() && left.endColumn() == right.endColumn();
    }

    private static boolean rangesOverlap(RangeRef left, RangeRef right) {
        return left.startRow() <= right.endRow() && right.startRow() <= left.endRow()
                && left.startColumn() <= right.endColumn() && right.startColumn() <= left.endColumn();
    }

    private RangeRef filterRange(ObjectNode root, String sheetId, ObjectNode params) {
        return SnapshotMutationSupport.range(root, filter(root, sheetId, params).get("range"));
    }

    private List<RangeRef> existingFilterRange(ObjectNode root, String sheetId) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        JsonNode filter = sheet.get("autoFilter");
        if (filter == null || filter.isNull()) return List.of();
        return List.of(SnapshotMutationSupport.range(root, SnapshotMutationSupport.requiredObject(sheet, "autoFilter").get("range")));
    }

    private void upsertRule(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params, String collection) {
        ObjectNode rule = SnapshotMutationSupport.requiredObject(params, "rule");
        validateRule(root, sheetId, rule);
        SnapshotMutationSupport.upsertById(SnapshotMutationSupport.array(sheet, collection), rule);
    }

    private void removeRule(ObjectNode sheet, ObjectNode params, String collection) {
        SnapshotMutationSupport.removeById(SnapshotMutationSupport.array(sheet, collection), SnapshotMutationSupport.text(params, "ruleId"));
    }

    private void validateRule(ObjectNode root, String sheetId, ObjectNode rule) {
        SnapshotMutationSupport.text(rule, "id");
        SnapshotMutationSupport.requireEntitySheet(rule, sheetId);
        ruleRanges(root, sheetId, rule);
    }

    private List<RangeRef> ruleRanges(ObjectNode root, String sheetId, ObjectNode rule) {
        return SnapshotMutationSupport.ranges(root, rule.get("ranges"), sheetId);
    }

    private List<RangeRef> existingRuleRanges(ObjectNode root, String sheetId, String property, String ruleId) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        ObjectNode rule = SnapshotMutationSupport.findById(SnapshotMutationSupport.array(sheet, property), ruleId);
        return rule == null ? List.of() : ruleRanges(root, sheetId, rule);
    }

    private List<RangeRef> allRuleRanges(ObjectNode root, String sheetId, String property) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        List<RangeRef> ranges = new ArrayList<>();
        for (JsonNode rule : SnapshotMutationSupport.array(sheet, property)) {
            if (!rule.isObject()) throw ServiceException.validation(property + " contains an invalid rule");
            ranges.addAll(ruleRanges(root, sheetId, (ObjectNode) rule));
        }
        return List.copyOf(ranges);
    }

    private void setBanded(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        JsonNode rule = params.get("rule");
        if (rule == null || rule.isNull()) {
            sheet.remove("bandedRule");
            return;
        }
        if (!rule.isObject()) throw ServiceException.validation("Banded rule must be an object or null");
        ObjectNode banded = (ObjectNode) rule;
        SnapshotMutationSupport.requireEntitySheet(banded, sheetId);
        SnapshotMutationSupport.range(root, banded.get("range"));
        sheet.set("bandedRule", banded.deepCopy());
    }

    private List<RangeRef> bandedRange(ObjectNode root, String sheetId, ObjectNode params) {
        JsonNode rule = params.get("rule");
        if (rule == null || rule.isNull()) return List.of();
        if (!rule.isObject()) throw ServiceException.validation("Banded rule must be an object or null");
        SnapshotMutationSupport.requireEntitySheet((ObjectNode) rule, sheetId);
        return List.of(SnapshotMutationSupport.range(root, rule.get("range")));
    }

    private void setOutline(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        ObjectNode outline = SnapshotMutationSupport.requiredObject(params, "outline");
        ArrayNode groups = SnapshotMutationSupport.requiredArray(outline, "groups");
        for (JsonNode group : groups) validateOutlineGroup(root, sheetId, group);
        sheet.set("outline", outline.deepCopy());
    }

    private void validateOutlineGroup(ObjectNode root, String sheetId, JsonNode group) {
        if (group == null || !group.isObject()) throw ServiceException.validation("Outline group must be an object");
        ObjectNode entry = (ObjectNode) group;
        SnapshotMutationSupport.text(entry, "id");
        String axis = SnapshotMutationSupport.text(entry, "axis");
        if (!axis.equals("row") && !axis.equals("column")) throw ServiceException.validation("Outline axis is invalid");
        String startKey = axis.equals("row") ? "row" : "column";
        ObjectNode start = JsonNodeFactory.instance.objectNode();
        ObjectNode end = JsonNodeFactory.instance.objectNode();
        start.set(startKey, entry.get("start"));
        end.set(startKey, entry.get("end"));
        int startValue = SnapshotMutationSupport.index(root, sheetId, start, startKey);
        int endValue = SnapshotMutationSupport.index(root, sheetId, end, startKey);
        if (endValue < startValue) throw ServiceException.validation("Outline bounds are invalid");
        JsonNode level = entry.get("level");
        if (level == null || !level.isIntegralNumber() || level.intValue() < 1 || level.intValue() > 3) throw ServiceException.validation("Outline level is invalid");
        if (!entry.path("collapsed").isBoolean()) throw ServiceException.validation("Outline collapsed must be boolean");
    }

    private RangeRef tableRange(ObjectNode root, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.requireEntitySheet(params, sheetId);
        SnapshotMutationSupport.text(params, "id");
        return ownRange(root, sheetId, params.get("range"));
    }

    private void upsertSheetTable(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        tableRange(root, sheetId, params);
        SnapshotMutationSupport.upsertById(SnapshotMutationSupport.array(sheet, "sheetTables"), params);
    }

    private void updateTableSheet(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.requireEntitySheet(params, sheetId);
        if (!"table-sheet".equals(sheet.path("kind").asText())) throw ServiceException.validation("TableSheet definition targets a non-TableSheet");
        ObjectNode definition = SnapshotMutationSupport.requiredObject(params, "definition");
        String viewId = SnapshotMutationSupport.text(definition, "viewId");
        ObjectNode table = null;
        for (JsonNode candidate : SnapshotMutationSupport.dataModelArray(root, "tables")) {
            if (viewId.equals(candidate.path("id").asText()) && candidate.isObject()) {
                table = (ObjectNode) candidate;
                break;
            }
        }
        if (table == null) throw ServiceException.validation("TableSheet binding table is unavailable: " + viewId);
        Set<String> fieldIds = new java.util.HashSet<>();
        for (JsonNode field : SnapshotMutationSupport.array(table, "fields")) {
            fieldIds.add(SnapshotMutationSupport.text((ObjectNode) field, "id"));
        }
        ArrayNode columns = SnapshotMutationSupport.requiredArray(definition, "columns");
        if (columns.isEmpty()) throw ServiceException.validation("TableSheet must expose at least one column");
        Set<String> visibleIds = new java.util.HashSet<>();
        for (JsonNode rawColumn : columns) {
            if (!rawColumn.isObject()) throw ServiceException.validation("TableSheet column is invalid");
            ObjectNode column = (ObjectNode) rawColumn;
            String fieldId = SnapshotMutationSupport.text(column, "fieldId");
            String caption = SnapshotMutationSupport.text(column, "caption");
            if (!fieldIds.contains(fieldId) || !visibleIds.add(fieldId) || caption.isBlank()) throw ServiceException.validation("TableSheet column does not match its binding table");
            JsonNode width = column.get("widthPx");
            if (width != null && (!width.isNumber() || width.asDouble() <= 0)) throw ServiceException.validation("TableSheet column width is invalid");
            JsonNode type = column.get("type");
            if (type != null && !type.isTextual()) throw ServiceException.validation("TableSheet column type is invalid");
            JsonNode formula = column.get("formula");
            if (formula != null && !formula.isTextual()) throw ServiceException.validation("TableSheet column formula is invalid");
        }
        validateTableSheetFieldList(definition, visibleIds, "grouping");
        JsonNode sortState = definition.get("sortState");
        if (sortState != null) {
            if (!sortState.isArray()) throw ServiceException.validation("TableSheet sortState is invalid");
            Set<String> sortedIds = new java.util.HashSet<>();
            for (JsonNode rawSort : sortState) {
                if (!rawSort.isObject()) throw ServiceException.validation("TableSheet sort state is invalid");
                String fieldId = SnapshotMutationSupport.text((ObjectNode) rawSort, "fieldId");
                String direction = SnapshotMutationSupport.text((ObjectNode) rawSort, "direction");
                if (!visibleIds.contains(fieldId) || !sortedIds.add(fieldId) || !(direction.equals("asc") || direction.equals("desc"))) throw ServiceException.validation("TableSheet sort state does not match its visible columns");
            }
        }
        sheet.set("tableSheet", definition.deepCopy());
    }

    private void updateGanttSheet(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.requireEntitySheet(params, sheetId);
        if (!"gantt-sheet".equals(sheet.path("kind").asText())) throw ServiceException.validation("GanttSheet definition targets a non-GanttSheet");
        ObjectNode definition = SnapshotMutationSupport.requiredObject(params, "definition");
        String viewId = SnapshotMutationSupport.text(definition, "viewId");
        ObjectNode table = null;
        for (JsonNode candidate : SnapshotMutationSupport.dataModelArray(root, "tables")) {
            if (viewId.equals(candidate.path("id").asText()) && candidate.isObject()) { table = (ObjectNode) candidate; break; }
        }
        if (table == null) throw ServiceException.validation("GanttSheet binding table is unavailable: " + viewId);
        Set<String> fieldIds = new java.util.HashSet<>();
        for (JsonNode field : SnapshotMutationSupport.array(table, "fields")) fieldIds.add(SnapshotMutationSupport.text((ObjectNode) field, "id"));
        ObjectNode mapping = SnapshotMutationSupport.requiredObject(definition, "fieldMap");
        for (String key : List.of("id", "title", "start", "end", "progress")) {
            String fieldId = SnapshotMutationSupport.text(mapping, key);
            if (!fieldIds.contains(fieldId)) throw ServiceException.validation("GanttSheet field mapping is unavailable: " + key);
        }
        for (String key : List.of("parentId", "dependencies")) {
            JsonNode field = mapping.get(key);
            if (field != null && !field.isNull() && (!field.isTextual() || !fieldIds.contains(field.asText()))) throw ServiceException.validation("GanttSheet optional field mapping is unavailable: " + key);
        }
        ObjectNode calendar = SnapshotMutationSupport.requiredObject(definition, "calendar");
        ArrayNode workingDays = SnapshotMutationSupport.requiredArray(calendar, "workingDays");
        Set<Integer> days = new java.util.HashSet<>();
        for (JsonNode day : workingDays) {
            if (!day.isIntegralNumber() || day.intValue() < 0 || day.intValue() > 6 || !days.add(day.intValue())) throw ServiceException.validation("GanttSheet workingDays is invalid");
        }
        JsonNode startHour = calendar.get("dayStartHour");
        JsonNode endHour = calendar.get("dayEndHour");
        if (startHour == null || endHour == null || !startHour.isNumber() || !endHour.isNumber() || startHour.asDouble() < 0 || endHour.asDouble() > 24 || startHour.asDouble() >= endHour.asDouble()) throw ServiceException.validation("GanttSheet calendar hours are invalid");
        ObjectNode timeline = SnapshotMutationSupport.requiredObject(definition, "timeline");
        String unit = SnapshotMutationSupport.text(timeline, "unit");
        if (!Set.of("day", "week", "month", "quarter").contains(unit)) throw ServiceException.validation("GanttSheet timeline unit is invalid");
        for (String key : List.of("start", "end")) if (timeline.get(key) != null && !timeline.get(key).isTextual()) throw ServiceException.validation("GanttSheet timeline bound is invalid");
        ObjectNode dependencyStyle = SnapshotMutationSupport.requiredObject(definition, "dependencyStyle");
        if (SnapshotMutationSupport.text(dependencyStyle, "color").isBlank() || !dependencyStyle.path("width").isNumber() || dependencyStyle.path("width").asDouble() <= 0) throw ServiceException.validation("GanttSheet dependency style is invalid");
        sheet.set("ganttSheet", definition.deepCopy());
    }

    private void validateTableSheetFieldList(ObjectNode definition, Set<String> visibleIds, String property) {
        ArrayNode entries = SnapshotMutationSupport.requiredArray(definition, property);
        Set<String> ids = new java.util.HashSet<>();
        for (JsonNode raw : entries) {
            if (!raw.isObject()) throw ServiceException.validation("TableSheet " + property + " entry is invalid");
            String fieldId = SnapshotMutationSupport.text((ObjectNode) raw, "fieldId");
            if (!visibleIds.contains(fieldId) || !ids.add(fieldId)) throw ServiceException.validation("TableSheet " + property + " does not match its visible columns");
            JsonNode collapsed = raw.get("collapsed");
            if (collapsed != null && !collapsed.isBoolean()) throw ServiceException.validation("TableSheet grouping collapsed flag is invalid");
        }
    }

    private RangeRef existingTableRange(ObjectNode root, String sheetId, String tableId) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        ObjectNode table = SnapshotMutationSupport.findById(SnapshotMutationSupport.array(sheet, "sheetTables"), tableId);
        return table == null ? SnapshotMutationSupport.wholeSheetRange(root, sheetId) : ownRange(root, sheetId, table.get("range"));
    }

    private void removeSheetTable(ObjectNode sheet, ObjectNode params) {
        SnapshotMutationSupport.removeById(SnapshotMutationSupport.array(sheet, "sheetTables"), SnapshotMutationSupport.text(params, "tableId"));
    }

    private RangeRef ownRange(ObjectNode root, String sheetId, JsonNode value) {
        RangeRef range = SnapshotMutationSupport.range(root, value);
        SnapshotMutationSupport.requireSheet(range, sheetId);
        return range;
    }

    private int sheetIndex(ArrayNode sheets, String sheetId) {
        for (int index = 0; index < sheets.size(); index++) if (sheetId.equals(sheets.get(index).path("id").asText())) return index;
        return -1;
    }

    private ObjectNode renameIndex(ObjectNode params, String source, String target) {
        JsonNode value = params.get(source);
        if (value == null) throw ServiceException.validation(source + " is required");
        ObjectNode result = JsonNodeFactory.instance.objectNode();
        result.set(target, value);
        return result;
    }

    private boolean containsIndex(ArrayNode values, int candidate) {
        for (JsonNode value : values) if (value.isIntegralNumber() && value.intValue() == candidate) return true;
        return false;
    }

    private static boolean checksProtection(String id) {
        return !id.equals("sheet.reordered");
    }

    private static String protectionAction(String id) {
        if (id.startsWith("sheetTable") || id.startsWith("tableSheet") || id.equals("sheet.reordered")) return "structure";
        if (id.startsWith("cf") || id.startsWith("dv") || id.equals("banded.set") || id.equals("outline.set")) return "format";
        return "edit-cell";
    }
}
