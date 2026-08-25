package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

/**
 * The single server-side protection authority.  The snapshot is the source
 * of truth: clients cannot widen a range, opt out of a rule, or replace the
 * native {@code allow} contract with an older action list.
 */
final class ProtectionResolver {
    private static final Set<String> OPERATION_ACTIONS = Set.of(
            "format", "insert-rows", "insert-columns", "delete-rows", "delete-columns",
            "sort", "auto-filter", "edit-objects", "select-locked", "select-unlocked"
    );

    private ProtectionResolver() {
    }

    static void assertAllowed(JsonNode snapshot, List<RangeRef> ranges, String action) {
        if (ranges == null || ranges.isEmpty()) return;
        if (action == null || action.isBlank()) throw ServiceException.validation("Protection action is required");
        if ("structure".equals(action)) action = "format";
        if ("drawing".equals(action)) action = "edit-objects";
        if ("print".equals(action)) return;
        if (!"edit-cell".equals(action) && !OPERATION_ACTIONS.contains(action)) {
            throw ServiceException.validation("Unknown protection action: " + action);
        }
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        for (RangeRef target : ranges) {
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, target.sheetId());
            if ("edit-cell".equals(action)) {
                assertEditableCells(root, sheet, target);
            } else {
                assertOperationAllowed(root, sheet, target, action);
            }
        }
    }

    private static void assertOperationAllowed(ObjectNode root, ObjectNode sheet, RangeRef target, String action) {
        String allowField = allowField(action);
        for (JsonNode rule : matchingRules(root, sheet, target)) {
            if (!rule.path("locked").isBoolean() || !rule.path("locked").asBoolean()) continue;
            JsonNode allow = rule.get("allow");
            if (allow == null || !allow.isObject() || !allow.path(allowField).isBoolean() || !allow.path(allowField).asBoolean()) {
                throw ServiceException.forbidden("Protected area blocks mutation " + action);
            }
        }
    }

    private static void assertEditableCells(ObjectNode root, ObjectNode sheet, RangeRef target) {
        int rowCount = boundedDimension(sheet, "rowCount", SnapshotMutationSupport.MAX_ROW + 1);
        int columnCount = boundedDimension(sheet, "columnCount", SnapshotMutationSupport.MAX_COLUMN + 1);
        if (target.startRow() >= rowCount || target.startColumn() >= columnCount) return;
        int endRow = Math.min(target.endRow(), rowCount - 1);
        int endColumn = Math.min(target.endColumn(), columnCount - 1);
        for (int row = target.startRow(); row <= endRow; row++) {
            for (int column = target.startColumn(); column <= endColumn; column++) {
                if (isLocked(root, sheet, target.sheetId(), row, column)) {
                    throw ServiceException.forbidden("Protected cell blocks mutation " + target.sheetId() + "!" + row + "," + column);
                }
            }
        }
    }

    private static boolean isLocked(ObjectNode root, ObjectNode sheet, String sheetId, int row, int column) {
        RangeRef cell = new RangeRef(sheetId, row, row, column, column);
        List<JsonNode> matches = matchingRules(root, sheet, cell);
        boolean active = matches.stream().anyMatch(rule -> rule.path("locked").asBoolean(false)
                && ("workbook".equals(rule.path("scope").asText()) || "sheet".equals(rule.path("scope").asText())));
        boolean rangeLocked = matches.stream().anyMatch(rule -> "range".equals(rule.path("scope").asText()) && rule.path("locked").asBoolean(false));
        if (rangeLocked) return true;
        if (!active) return false;
        JsonNode cells = sheet.get("cells");
        JsonNode cellNode = cells != null && cells.isObject()
                ? cells.path(Integer.toString(row)).path(Integer.toString(column))
                : null;
        if (cellNode != null && cellNode.isMissingNode()) cellNode = null;
        JsonNode style = cellNode == null ? null : cellNode.get("style");
        return style == null || !style.isObject() || !style.path("locked").isBoolean() || style.path("locked").asBoolean();
    }

    private static List<JsonNode> matchingRules(ObjectNode root, ObjectNode sheet, RangeRef target) {
        List<JsonNode> result = new ArrayList<>();
        for (JsonNode candidateSheet : SnapshotMutationSupport.sheets(root)) {
            if (!candidateSheet.isObject()) continue;
            boolean sameSheet = target.sheetId().equals(candidateSheet.path("id").asText());
            JsonNode rules = candidateSheet.get("protectionRules");
            if (rules == null || !rules.isArray()) continue;
            for (JsonNode rule : rules) {
                if (!rule.isObject() || !rule.path("locked").isBoolean()) continue;
                String scope = rule.path("scope").asText();
                if ("workbook".equals(scope) || (sameSheet && "sheet".equals(scope))) {
                    result.add(rule);
                } else if (sameSheet && "range".equals(scope) && intersects(rule.get("range"), target)) {
                    result.add(rule);
                }
            }
        }
        return result;
    }

    private static boolean intersects(JsonNode raw, RangeRef target) {
        if (raw == null || !raw.isObject()) return false;
        if (!target.sheetId().equals(raw.path("sheetId").asText())) return false;
        return target.startRow() <= raw.path("endRow").asInt(-1)
                && target.endRow() >= raw.path("startRow").asInt(Integer.MAX_VALUE)
                && target.startColumn() <= raw.path("endColumn").asInt(-1)
                && target.endColumn() >= raw.path("startColumn").asInt(Integer.MAX_VALUE);
    }

    private static String allowField(String action) {
        return switch (action) {
            case "format" -> "formatCells";
            case "insert-rows" -> "insertRows";
            case "insert-columns" -> "insertColumns";
            case "delete-rows" -> "deleteRows";
            case "delete-columns" -> "deleteColumns";
            case "sort" -> "sort";
            case "auto-filter" -> "autoFilter";
            case "edit-objects" -> "editObjects";
            case "select-locked" -> "selectLocked";
            case "select-unlocked" -> "selectUnlocked";
            default -> throw ServiceException.validation("Protection action has no allow field: " + action);
        };
    }

    private static int boundedDimension(ObjectNode sheet, String field, int fallback) {
        JsonNode value = sheet.get(field);
        if (value == null || !value.isIntegralNumber() || value.intValue() < 1) return fallback;
        return Math.min(value.intValue(), fallback);
    }
}
