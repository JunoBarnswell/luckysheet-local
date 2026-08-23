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
            "filter.set", "filter.remove",
            "cf.add", "cf.remove", "cf.clear",
            "dv.add", "dv.remove", "banded.set", "outline.set",
            "sheetTable.add", "sheetTable.remove", "sheetTable.update"
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
            case "filter.set" -> List.of(filterRange(root, mutation.sheetId(), params));
            case "filter.remove" -> existingFilterRange(root, mutation.sheetId());
            case "cf.add" -> ruleRanges(root, mutation.sheetId(), SnapshotMutationSupport.requiredObject(params, "rule"));
            case "cf.remove" -> existingRuleRanges(root, mutation.sheetId(), "conditionalFormats", SnapshotMutationSupport.text(params, "ruleId"));
            case "cf.clear" -> allRuleRanges(root, mutation.sheetId(), "conditionalFormats");
            case "dv.add" -> ruleRanges(root, mutation.sheetId(), SnapshotMutationSupport.requiredObject(params, "rule"));
            case "dv.remove" -> existingRuleRanges(root, mutation.sheetId(), "dataValidations", SnapshotMutationSupport.text(params, "ruleId"));
            case "banded.set" -> bandedRange(root, mutation.sheetId(), params);
            case "outline.set" -> List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
            case "sheetTable.add", "sheetTable.update" -> List.of(tableRange(root, mutation.sheetId(), params));
            case "sheetTable.remove" -> List.of(existingTableRange(root, mutation.sheetId(), SnapshotMutationSupport.text(params, "tableId")));
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
            case "filter.set" -> sheet.set("filter", filter(root, mutation.sheetId(), params).deepCopy());
            case "filter.remove" -> sheet.remove("filter");
            case "cf.add" -> upsertRule(root, sheet, mutation.sheetId(), params, "conditionalFormats");
            case "cf.remove" -> removeRule(sheet, params, "conditionalFormats");
            case "cf.clear" -> sheet.set("conditionalFormats", JsonNodeFactory.instance.arrayNode());
            case "dv.add" -> upsertRule(root, sheet, mutation.sheetId(), params, "dataValidations");
            case "dv.remove" -> removeRule(sheet, params, "dataValidations");
            case "banded.set" -> setBanded(root, sheet, mutation.sheetId(), params);
            case "outline.set" -> setOutline(root, sheet, mutation.sheetId(), params);
            case "sheetTable.add", "sheetTable.update" -> upsertSheetTable(root, sheet, mutation.sheetId(), params);
            case "sheetTable.remove" -> removeSheetTable(sheet, params);
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
        ObjectNode filter = SnapshotMutationSupport.requiredObject(params, "filter");
        SnapshotMutationSupport.requireEntitySheet(filter, sheetId);
        RangeRef range = SnapshotMutationSupport.range(root, filter.get("range"));
        SnapshotMutationSupport.requireSheet(range, sheetId);
        JsonNode criteria = filter.get("criteria");
        if (criteria == null || !criteria.isObject()) throw ServiceException.validation("Filter criteria must be an object");
        return filter;
    }

    private RangeRef filterRange(ObjectNode root, String sheetId, ObjectNode params) {
        return SnapshotMutationSupport.range(root, filter(root, sheetId, params).get("range"));
    }

    private List<RangeRef> existingFilterRange(ObjectNode root, String sheetId) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        JsonNode filter = sheet.get("filter");
        if (filter == null || filter.isNull()) return List.of();
        return List.of(SnapshotMutationSupport.range(root, SnapshotMutationSupport.requiredObject(sheet, "filter").get("range")));
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
        if (id.startsWith("sheetTable") || id.equals("sheet.reordered")) return "structure";
        if (id.startsWith("cf") || id.startsWith("dv") || id.equals("banded.set") || id.equals("outline.set")) return "format";
        return "edit-cell";
    }
}
