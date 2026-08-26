package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Server-side authority for the semantic source of a cell write.
 *
 * A browser-side validation result is evidence only.  Direct entry is
 * evaluated again against the snapshot that is about to be committed, while
 * paste/fill/formula/script writes retain their derived-write semantics.
 */
final class CellEntryIntentAuthority {
    private static final List<String> KINDS = List.of(
            "direct-entry", "paste", "fill", "formula-result", "query-load", "script", "external-sync", "restore");

    private CellEntryIntentAuthority() {
    }

    static void requireCellWrite(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params, ObjectNode value) {
        ObjectNode intent = requireIntent(params);
        String kind = requiredText(intent, "kind", "entryIntent");
        if (!KINDS.contains(kind)) throw ServiceException.validation("entryIntent.kind is unsupported: " + kind);
        if ("direct-entry".equals(kind)) {
            requireDirectCandidate(intent, sheetId, params, value);
            validateDirectEntry(root, sheet, sheetId, params, value, intent);
        }
    }

    static void requireRangeWrite(ObjectNode params, String mutationId) {
        ObjectNode intent = requireIntent(params);
        String kind = requiredText(intent, "kind", "entryIntent");
        if (!KINDS.contains(kind)) throw ServiceException.validation("entryIntent.kind is unsupported: " + kind);
        if ("direct-entry".equals(kind)) {
            throw ServiceException.validation("RANGE_DIRECT_ENTRY_FORBIDDEN: direct entry must use cell.set");
        }
        if ("range.paste".equals(mutationId) && !"paste".equals(kind)) {
            throw ServiceException.validation("range.paste requires entryIntent.kind=paste");
        }
    }

    private static ObjectNode requireIntent(ObjectNode params) {
        JsonNode value = params.get("entryIntent");
        if (value == null || !value.isObject()) throw ServiceException.validation("entryIntent is required for cell writes");
        ObjectNode intent = (ObjectNode) value;
        JsonNode decision = intent.get("validationDecision");
        if (decision == null || !decision.isObject()) throw ServiceException.validation("entryIntent.validationDecision is required");
        String status = requiredText((ObjectNode) decision, "status", "entryIntent.validationDecision");
        if (!List.of("accepted", "confirmed", "rejected", "not-applicable").contains(status)) {
            throw ServiceException.validation("entryIntent.validationDecision.status is unsupported: " + status);
        }
        return intent;
    }

    private static void requireDirectCandidate(ObjectNode intent, String sheetId, ObjectNode params, ObjectNode value) {
        JsonNode target = intent.get("target");
        if (target == null || !target.isObject()
                || !sheetId.equals(target.path("sheetId").asText())
                || target.path("row").asInt(-1) != params.path("row").asInt(-2)
                || target.path("column").asInt(-1) != params.path("column").asInt(-2)) {
            throw ServiceException.validation("entryIntent.target does not match cell.set target");
        }
        JsonNode candidate = intent.get("candidate");
        if (candidate == null || !candidate.isObject() || !candidate.equals(value)) {
            throw ServiceException.validation("entryIntent.candidate does not match cell.set value");
        }
    }

    private static void validateDirectEntry(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params, ObjectNode value, ObjectNode intent) {
        JsonNode formula = value.get("formula");
        if (formula != null && !formula.isNull()) {
            requireDecision(intent, true, null, null);
            return;
        }
        int row = params.path("row").asInt(-1);
        int column = params.path("column").asInt(-1);
        ObjectNode rule = findRule(root, sheet, sheetId, row, column);
        if (rule == null) {
            requireDecision(intent, true, null, null);
            return;
        }
        String ruleId = textOrNull(rule, "id");
        JsonNode declaredRuleId = intent.path("validationDecision").get("ruleId");
        if (declaredRuleId != null && !declaredRuleId.isNull() && !declaredRuleId.isTextual()) {
            throw ServiceException.validation("entryIntent.validationDecision.ruleId must be a string");
        }
        if (declaredRuleId != null && !declaredRuleId.isNull() && !declaredRuleId.asText().equals(ruleId)) {
            throw ServiceException.validation("CELL_ENTRY_VALIDATION_STALE: validation rule changed before commit");
        }
        ValidationResult result = evaluateRule(root, sheet, sheetId, row, column, value, rule);
        requireDecision(intent, result.valid(), result.alertStyle(), result.message());
    }

    private static ObjectNode findRule(ObjectNode root, ObjectNode sheet, String sheetId, int row, int column) {
        JsonNode rules = sheet.get("dataValidations");
        if (rules == null || !rules.isArray()) return null;
        for (JsonNode ruleNode : rules) {
            if (!ruleNode.isObject()) throw ServiceException.validation("dataValidations contains a non-object rule");
            JsonNode ranges = ruleNode.get("ranges");
            if (ranges == null || !ranges.isArray()) throw ServiceException.validation("data validation ranges are required");
            for (JsonNode rangeNode : ranges) {
                RangeRef range = SnapshotMutationSupport.range(root, rangeNode);
                if (sheetId.equals(range.sheetId()) && row >= range.startRow() && row <= range.endRow()
                        && column >= range.startColumn() && column <= range.endColumn()) return (ObjectNode) ruleNode;
            }
        }
        return null;
    }

    private static ValidationResult evaluateRule(ObjectNode root, ObjectNode sheet, String sheetId, int row, int column, ObjectNode value, ObjectNode rule) {
        String type = requiredText(rule, "type", "data validation rule").toLowerCase(Locale.ROOT);
        String alertStyle = rule.path("alertStyle").asText("stop").toLowerCase(Locale.ROOT);
        if (!List.of("stop", "warning", "information").contains(alertStyle)) throw ServiceException.validation("Unsupported data validation alertStyle: " + alertStyle);
        JsonNode candidate = value.get("value");
        boolean blank = candidate == null || candidate.isNull() || (candidate.isTextual() && candidate.asText().isEmpty());
        if (blank) return new ValidationResult(rule.path("allowBlank").asBoolean(true), alertStyle, "该单元格不允许为空");
        if ("list".equals(type)) {
            List<String> allowed = listValues(root, sheet, sheetId, rule);
            boolean valid = allowed.stream().anyMatch(item -> item.equalsIgnoreCase(candidate.asText()));
            return new ValidationResult(valid, alertStyle, "值不在允许的列表中");
        }
        if ("checkbox".equals(type)) {
            boolean valid = candidate.isBoolean() || (candidate.isTextual() && (candidate.asText().equalsIgnoreCase("true") || candidate.asText().equalsIgnoreCase("false")));
            return new ValidationResult(valid, alertStyle, "需要 TRUE/FALSE");
        }
        if ("whole".equals(type) || "decimal".equals(type)) {
            if (!candidate.isNumber()) return new ValidationResult(false, alertStyle, "需要输入数字");
            double actual = candidate.asDouble();
            if ("whole".equals(type) && actual != Math.rint(actual)) return new ValidationResult(false, alertStyle, "需要输入整数");
            return compareNumeric(actual, rule, alertStyle);
        }
        if ("textlength".equals(type)) {
            return compareNumeric(candidate.asText().length(), rule, alertStyle);
        }
        if ("date".equals(type)) {
            if (!candidate.isNumber()) return new ValidationResult(false, alertStyle, "需要输入有效日期");
            return compareNumeric(candidate.asDouble(), rule, alertStyle);
        }
        if ("time".equals(type)) {
            if (!candidate.isNumber() || candidate.asDouble() < 0 || candidate.asDouble() >= 1) return new ValidationResult(false, alertStyle, "需要输入有效时间");
            return compareNumeric(candidate.asDouble(), rule, alertStyle);
        }
        if ("custom".equals(type)) throw ServiceException.unavailable("UNSUPPORTED_FEATURE: server cannot authoritatively evaluate custom data validation formulas");
        throw ServiceException.unavailable("UNSUPPORTED_FEATURE: server cannot authoritatively evaluate data validation type " + type);
    }

    private static List<String> listValues(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode rule) {
        JsonNode source = rule.get("listSource");
        if (source != null && source.isObject()) {
            String kind = source.path("kind").asText();
            if ("values".equals(kind)) {
                JsonNode values = source.get("values");
                if (values == null || !values.isArray()) throw ServiceException.validation("list validation values are required");
                List<String> result = new ArrayList<>();
                values.forEach(value -> result.add(value.asText()));
                return result;
            }
            if ("range".equals(kind)) {
                RangeRef range = SnapshotMutationSupport.range(root, source.get("range"));
                if (!sheetId.equals(range.sheetId())) throw ServiceException.validation("list validation range must target the same sheet");
                ObjectNode sourceSheet = SnapshotMutationSupport.sheet(root, sheetId);
                List<String> result = new ArrayList<>();
                for (int row = range.startRow(); row <= range.endRow(); row++) for (int column = range.startColumn(); column <= range.endColumn(); column++) {
                    JsonNode cell = SnapshotMutationSupport.cell(sourceSheet, new SnapshotMutationSupport.CellCoordinate(row, column), false);
                    if (cell != null && cell.has("value") && !cell.get("value").isNull()) result.add(cell.get("value").asText());
                }
                return result;
            }
            throw ServiceException.unavailable("UNSUPPORTED_FEATURE: formula-backed list validation requires shared evaluator authority");
        }
        String formula = textOrNull(rule, "formula1");
        if (formula == null) throw ServiceException.validation("list validation source is required");
        List<String> result = new ArrayList<>();
        for (String item : formula.replaceFirst("^=", "").split(",", -1)) if (!item.isBlank()) result.add(item.trim().replace("\"", ""));
        return result;
    }

    private static ValidationResult compareNumeric(double actual, ObjectNode rule, String alertStyle) {
        String operator = rule.path("operator").asText("between");
        Double first = numberOrNull(rule.get("formula1"));
        Double second = numberOrNull(rule.get("formula2"));
        boolean valid = switch (operator) {
            case "greaterThan" -> first != null && actual > first;
            case "lessThan" -> first != null && actual < first;
            case "equal" -> first != null && actual == first;
            case "notEqual" -> first != null && actual != first;
            case "notBetween" -> first != null && second != null && (actual < first || actual > second);
            case "between" -> first == null || second == null || (actual >= first && actual <= second);
            default -> throw ServiceException.validation("Unsupported data validation operator: " + operator);
        };
        return new ValidationResult(valid, alertStyle, "不符合数据验证规则");
    }

    private static void requireDecision(ObjectNode intent, boolean valid, String alertStyle, String message) {
        ObjectNode decision = (ObjectNode) intent.get("validationDecision");
        String status = decision.path("status").asText();
        JsonNode declaredAlertStyle = decision.get("alertStyle");
        if (declaredAlertStyle != null && !declaredAlertStyle.isNull()
                && (alertStyle == null || !declaredAlertStyle.asText().equals(alertStyle))) {
            throw ServiceException.validation("CELL_ENTRY_VALIDATION_STALE: validation alert style changed before commit");
        }
        if (valid) {
            if (!List.of("accepted", "confirmed").contains(status)) throw ServiceException.validation("CELL_ENTRY_DECISION_INVALID: accepted direct entry decision is required");
            return;
        }
        if ("stop".equals(alertStyle)) throw ServiceException.validation(message == null ? "Cell value failed data validation" : message);
        if (!"confirmed".equals(status)) throw ServiceException.validation("CELL_ENTRY_CONFIRMATION_REQUIRED: warning/information validation requires explicit confirmation");
    }

    private static Double numberOrNull(JsonNode value) {
        if (value == null || value.isNull()) return null;
        if (value.isNumber()) return value.asDouble();
        if (!value.isTextual()) return null;
        try { return Double.valueOf(value.asText()); } catch (NumberFormatException ignored) { return null; }
    }

    private static String requiredText(ObjectNode object, String field, String label) {
        JsonNode value = object.get(field);
        if (value == null || !value.isTextual() || value.asText().isBlank()) throw ServiceException.validation(label + "." + field + " is required");
        return value.asText();
    }

    private static String textOrNull(ObjectNode object, String field) {
        JsonNode value = object.get(field);
        return value == null || value.isNull() ? null : value.isTextual() ? value.asText() : null;
    }

    private record ValidationResult(boolean valid, String alertStyle, String message) {
    }
}
