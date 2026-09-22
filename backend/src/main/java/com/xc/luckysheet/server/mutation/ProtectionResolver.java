package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.GeneratedWorkbookContract;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;

/**
 * The single server-side protection authority.  The snapshot is the source
 * of truth: clients cannot widen a range, opt out of a rule, or replace the
 * native {@code allow} contract with an older action list.
 */
final class ProtectionResolver {
    private ProtectionResolver() {
    }

    static void assertAllowed(JsonNode snapshot, List<RangeRef> ranges, String action) {
        if (ranges == null || ranges.isEmpty()) return;
        if (action == null || action.isBlank()) throw ServiceException.validation("Protection action is required");
        if (!"edit-cell".equals(action) && GeneratedWorkbookContract.protectionAllowField(action) == null) {
            throw ServiceException.validation("Unknown protection action: " + action);
        }
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        for (RangeRef target : ranges) {
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, target.sheetId());
            assertCanonicalRange(sheet, target);
            if ("edit-cell".equals(action)) {
                assertEditableCells(root, sheet, target);
            } else {
                assertOperationAllowed(root, sheet, target, action);
            }
        }
    }

    private static void assertOperationAllowed(ObjectNode root, ObjectNode sheet, RangeRef target, String action) {
        String allowField = GeneratedWorkbookContract.protectionAllowField(action);
        if (allowField == null) throw ServiceException.validation("Protection action has no allow field: " + action);
        for (JsonNode rule : matchingRules(root, sheet, target)) {
            if (!rule.path("locked").isBoolean() || !rule.path("locked").asBoolean()) continue;
            JsonNode allow = rule.get("allow");
            if (allow == null || !allow.isObject() || !allow.path(allowField).isBoolean() || !allow.path(allowField).asBoolean()) {
                throw ServiceException.forbidden("Protected area blocks mutation " + action);
            }
        }
    }

    private static void assertEditableCells(ObjectNode root, ObjectNode sheet, RangeRef target) {
        List<JsonNode> matches = matchingRules(root, sheet, target);
        for (JsonNode rule : matches) {
            if (rule.path("locked").asBoolean(false) && "range".equals(rule.path("scope").asText())) {
                throw ServiceException.forbidden("Protected range blocks mutation " + target.sheetId());
            }
        }
        boolean active = matches.stream().anyMatch(rule -> rule.path("locked").asBoolean(false)
                && ("workbook".equals(rule.path("scope").asText()) || "sheet".equals(rule.path("scope").asText())));
        if (!active) return;
        long requestedCells = SnapshotMutationSupport.cellCount(target);
        long explicitUnlocked = countExplicitUnlockedCells(sheet, target);
        if (explicitUnlocked != requestedCells) {
            throw ServiceException.forbidden("Protected cell blocks mutation " + target.sheetId() + "!" + target.startRow() + "," + target.startColumn());
        }
    }

    private static void assertCanonicalRange(ObjectNode sheet, RangeRef target) {
        int rowCount = canonicalDimension(sheet, "rowCount");
        int columnCount = canonicalDimension(sheet, "columnCount");
        if (target.endRow() >= rowCount || target.endColumn() >= columnCount) {
            throw ServiceException.validation("Protection range exceeds canonical worksheet bounds");
        }
    }

    private static int canonicalDimension(ObjectNode sheet, String field) {
        JsonNode value = sheet.get(field);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 1) {
            throw ServiceException.validation("Protection requires canonical worksheet " + field);
        }
        return value.intValue();
    }

    private static long countExplicitUnlockedCells(ObjectNode sheet, RangeRef target) {
        JsonNode cells = sheet.get("cells");
        if (cells == null || !cells.isObject()) return 0;
        long count = 0;
        var rows = cells.fields();
        while (rows.hasNext()) {
            var rowEntry = rows.next();
            int row;
            try {
                row = Integer.parseInt(rowEntry.getKey());
            } catch (NumberFormatException ignored) {
                throw ServiceException.validation("Protection cell index contains an invalid row");
            }
            if (row < target.startRow() || row > target.endRow() || !rowEntry.getValue().isObject()) continue;
            var columns = rowEntry.getValue().fields();
            while (columns.hasNext()) {
                var columnEntry = columns.next();
                int column;
                try {
                    column = Integer.parseInt(columnEntry.getKey());
                } catch (NumberFormatException ignored) {
                    throw ServiceException.validation("Protection cell index contains an invalid column");
                }
                if (column < target.startColumn() || column > target.endColumn()) continue;
                JsonNode style = columnEntry.getValue().get("style");
                if (style != null && style.isObject() && style.path("locked").isBoolean() && !style.path("locked").asBoolean()) count += 1;
            }
        }
        return count;
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
                if ("workbook".equals(scope) || (sameSheet && "sheet".equals(scope)
                        && target.sheetId().equals(rule.path("sheetId").asText()))) {
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

}
