package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.function.Function;
import java.util.function.IntUnaryOperator;

/** Canonical server-side lifecycle for worksheet conditional-format and validation ranges. */
final class SheetRuleLifecycle {
    private SheetRuleLifecycle() {
    }

    static void crop(ObjectNode root, ObjectNode sheet, String sheetId, String property, RangeRef clear) {
        JsonNode existing = sheet.get(property);
        if (existing == null || existing.isNull()) return;
        if (!existing.isArray()) throw ServiceException.validation(property + " must be an array");
        ArrayNode nextRules = JsonNodeFactory.instance.arrayNode();
        for (JsonNode rule : existing) {
            if (!rule.isObject()) throw ServiceException.validation(property + " rule must be an object");
            ObjectNode copy = ((ObjectNode) rule).deepCopy();
            JsonNode ranges = copy.get("ranges");
            if (ranges == null || !ranges.isArray() || ranges.isEmpty()) throw ServiceException.validation(property + " rule ranges must be a non-empty array");
            ArrayNode remaining = JsonNodeFactory.instance.arrayNode();
            for (JsonNode candidate : ranges) {
                RangeRef source = SnapshotMutationSupport.range(root, candidate);
                if (!sheetId.equals(source.sheetId())) throw ServiceException.validation(property + " rule targets another sheet");
                for (RangeRef part : subtract(source, clear)) remaining.add(rangeNode(part));
            }
            if (!remaining.isEmpty()) {
                copy.set("ranges", remaining);
                nextRules.add(copy);
            }
        }
        sheet.set(property, nextRules);
    }

    static void validateRule(ObjectNode root, String sheetId, ObjectNode rule, String property) {
        SnapshotMutationSupport.text(rule, "id");
        SnapshotMutationSupport.requireEntitySheet(rule, sheetId);
        SnapshotMutationSupport.ranges(root, rule.get("ranges"), sheetId);
        JsonNode anchor = rule.get("formulaAnchor");
        if (anchor != null && !anchor.isNull()) {
            if (!anchor.isObject() || !sheetId.equals(anchor.path("sheetId").asText())
                    || !anchor.path("row").canConvertToInt() || !anchor.path("column").canConvertToInt()
                    || anchor.path("row").asInt(-1) < 0 || anchor.path("column").asInt(-1) < 0) {
                throw ServiceException.validation(property + " formulaAnchor is invalid");
            }
        }
        if (!"dataValidations".equals(property)) return;
        if ("list".equals(rule.path("type").asText()) && !rule.has("formula1") && !rule.has("listSource")) {
            throw ServiceException.validation("List validation requires a list source");
        }
        JsonNode listSource = rule.get("listSource");
        if (listSource == null || listSource.isNull()) return;
        if (!listSource.isObject() || !Set.of("range", "formula").contains(listSource.path("kind").asText())) {
            throw ServiceException.validation("Data validation list source is invalid");
        }
        if ("range".equals(listSource.path("kind").asText())) {
            RangeRef source = SnapshotMutationSupport.range(root, listSource.get("range"));
            SnapshotMutationSupport.requireSheet(source, sheetId);
        } else if (!listSource.path("formula").isTextual() || listSource.path("formula").asText().isBlank()) {
            throw ServiceException.validation("Data validation list formula is required");
        }
    }

    static void validateSnapshot(ObjectNode root, ObjectNode sheet, String sheetId, String property, JsonNode proposed, List<RangeRef> allowedRanges) {
        if (proposed == null || !proposed.isArray()) throw ServiceException.validation("Paste snapshot " + property + " must be an array");
        JsonNode existing = sheet.get(property);
        ArrayNode current = existing != null && existing.isArray() ? (ArrayNode) existing : JsonNodeFactory.instance.arrayNode();
        for (JsonNode rule : current) {
            if (!ownerIsContained(root, sheetId, rule, allowedRanges) && !containsJson(proposed, rule)) {
                throw ServiceException.validation("Paste snapshot changes an unrelated " + property + " rule");
            }
        }
        for (JsonNode rule : proposed) {
            if (!ownerIsContained(root, sheetId, rule, allowedRanges) && !containsJson(current, rule)) {
                throw ServiceException.validation("Paste snapshot adds an unrelated " + property + " rule");
            }
        }
    }

    static int affectedColumnEnd(ObjectNode root, ObjectNode sheet, int baseline) {
        int end = Math.max(baseline, sheet.path("columnCount").asInt(1) - 1);
        for (String property : List.of("conditionalFormats", "dataValidations")) {
            JsonNode rules = sheet.get(property);
            if (rules == null || rules.isNull()) continue;
            if (!rules.isArray()) throw ServiceException.validation(property + " must be an array");
            for (JsonNode raw : rules) {
                ObjectNode rule = requireRule(raw, property + " rule");
                JsonNode ranges = rule.get("ranges");
                if (ranges == null || !ranges.isArray()) throw ServiceException.validation(property + " rule ranges must be an array");
                for (JsonNode candidate : ranges) end = Math.max(end, SnapshotMutationSupport.range(root, candidate).endColumn());
            }
        }
        JsonNode protections = sheet.get("protectionRules");
        if (protections != null && protections.isArray()) {
            for (JsonNode raw : protections) if (raw.isObject() && raw.has("range")) end = Math.max(end, SnapshotMutationSupport.range(root, raw.get("range")).endColumn());
        }
        return end;
    }

    static void transformStructuralFields(ObjectNode root, ObjectNode sheet, String sheetId, RangeRef scope,
                                          Function<RangeRef, List<RangeRef>> mapRange, IntUnaryOperator mapRow) {
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "conditionalFormats")) {
            transformFormulaAnchor(root, requireRule(raw, "Conditional format"), sheetId, scope, mapRow);
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "dataValidations")) {
            ObjectNode rule = requireRule(raw, "Data validation");
            transformFormulaAnchor(root, rule, sheetId, scope, mapRow);
            JsonNode listSource = rule.get("listSource");
            if (listSource == null || listSource.isNull() || !listSource.isObject() || !"range".equals(listSource.path("kind").asText())) continue;
            RangeRef source = SnapshotMutationSupport.range(root, listSource.get("range"));
            List<RangeRef> mapped = mapRange.apply(source);
            if (mapped.size() != 1) throw ServiceException.validation("Row permutation cannot exactly remap validation list source");
            ((ObjectNode) listSource).set("range", rangeNode(mapped.get(0)));
        }
    }

    static void validateStructuralFields(ObjectNode root, ObjectNode sheet, String sheetId, RangeRef scope,
                                         Function<RangeRef, List<RangeRef>> mapRange) {
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "conditionalFormats")) {
            validateFormulaAnchor(root, requireRule(raw, "Conditional format"), sheetId, scope);
        }
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "dataValidations")) {
            ObjectNode rule = requireRule(raw, "Data validation");
            validateFormulaAnchor(root, rule, sheetId, scope);
            JsonNode listSource = rule.get("listSource");
            if (listSource == null || listSource.isNull() || !listSource.isObject() || !"range".equals(listSource.path("kind").asText())) continue;
            RangeRef source = SnapshotMutationSupport.range(root, listSource.get("range"));
            if (mapRange.apply(source).size() != 1) throw ServiceException.validation("Row permutation cannot exactly remap validation list source");
        }
    }

    private static ObjectNode requireRule(JsonNode raw, String label) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation(label + " must be an object");
        return (ObjectNode) raw;
    }

    private static void transformFormulaAnchor(ObjectNode root, ObjectNode rule, String sheetId, RangeRef scope, IntUnaryOperator mapRow) {
        JsonNode raw = rule.get("formulaAnchor");
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject() || !sheetId.equals(raw.path("sheetId").asText())
                || !raw.path("row").canConvertToInt() || !raw.path("column").canConvertToInt()
                || raw.path("row").asInt(-1) < 0 || raw.path("column").asInt(-1) < 0) {
            throw ServiceException.validation("Sheet rule formula anchor is invalid");
        }
        if (contains(scope, raw.path("row").asInt(), raw.path("column").asInt())) {
            ((ObjectNode) raw).put("row", mapRow.applyAsInt(raw.path("row").asInt()));
        }
    }

    private static void validateFormulaAnchor(ObjectNode root, ObjectNode rule, String sheetId, RangeRef scope) {
        JsonNode raw = rule.get("formulaAnchor");
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject() || !sheetId.equals(raw.path("sheetId").asText())
                || !raw.path("row").canConvertToInt() || !raw.path("column").canConvertToInt()
                || raw.path("row").asInt(-1) < 0 || raw.path("column").asInt(-1) < 0) {
            throw ServiceException.validation("Sheet rule formula anchor is invalid");
        }
    }

    private static boolean ownerIsContained(ObjectNode root, String sheetId, JsonNode rule, List<RangeRef> allowedRanges) {
        if (!rule.isObject()) throw ServiceException.validation("Paste owner rule must be an object");
        JsonNode ranges = rule.get("ranges");
        if (ranges == null || !ranges.isArray() || ranges.isEmpty()) throw ServiceException.validation("Paste owner rule ranges are required");
        for (JsonNode value : ranges) {
            RangeRef range = SnapshotMutationSupport.range(root, value);
            if (!sheetId.equals(range.sheetId())) throw ServiceException.validation("Paste owner rule targets another sheet");
            if (allowedRanges.stream().noneMatch(allowed -> allowed.sheetId().equals(range.sheetId())
                    && allowed.startRow() <= range.startRow() && allowed.endRow() >= range.endRow()
                    && allowed.startColumn() <= range.startColumn() && allowed.endColumn() >= range.endColumn())) return false;
        }
        return true;
    }

    private static boolean containsJson(JsonNode array, JsonNode candidate) {
        if (array == null || !array.isArray()) return false;
        for (JsonNode value : array) if (value.equals(candidate)) return true;
        return false;
    }

    private static List<RangeRef> subtract(RangeRef source, RangeRef clear) {
        if (source.sheetId().equals(clear.sheetId())
                && source.startRow() <= clear.endRow() && clear.startRow() <= source.endRow()
                && source.startColumn() <= clear.endColumn() && clear.startColumn() <= source.endColumn()) {
            int top = Math.max(source.startRow(), clear.startRow());
            int bottom = Math.min(source.endRow(), clear.endRow());
            int left = Math.max(source.startColumn(), clear.startColumn());
            int right = Math.min(source.endColumn(), clear.endColumn());
            List<RangeRef> result = new ArrayList<>();
            if (source.startRow() < top) result.add(new RangeRef(source.sheetId(), source.startRow(), top - 1, source.startColumn(), source.endColumn()));
            if (bottom < source.endRow()) result.add(new RangeRef(source.sheetId(), bottom + 1, source.endRow(), source.startColumn(), source.endColumn()));
            if (source.startColumn() < left) result.add(new RangeRef(source.sheetId(), top, bottom, source.startColumn(), left - 1));
            if (right < source.endColumn()) result.add(new RangeRef(source.sheetId(), top, bottom, right + 1, source.endColumn()));
            return result;
        }
        return List.of(source);
    }

    private static ObjectNode rangeNode(RangeRef range) {
        ObjectNode node = JsonNodeFactory.instance.objectNode();
        node.put("sheetId", range.sheetId());
        node.put("startRow", range.startRow());
        node.put("endRow", range.endRow());
        node.put("startColumn", range.startColumn());
        node.put("endColumn", range.endColumn());
        return node;
    }

    private static boolean contains(RangeRef range, int row, int column) {
        return range.startRow() <= row && row <= range.endRow()
                && range.startColumn() <= column && column <= range.endColumn();
    }
}
