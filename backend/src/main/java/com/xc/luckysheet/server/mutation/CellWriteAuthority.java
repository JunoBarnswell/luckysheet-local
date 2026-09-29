package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * Server-side authority for every cell.set write. Browser-side validation is
 * evidence only; the candidate is always re-evaluated against the snapshot
 * that is about to be committed.
 */
final class CellWriteAuthority {
    private static final List<String> KINDS = List.of(
            "direct-entry", "paste", "fill", "formula-result", "query-load", "script", "external-sync");

    private CellWriteAuthority() {
    }

    static void requireCellWrite(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params, ObjectNode value) {
        ObjectNode intent = requireIntent(params);
        String kind = requiredText(intent, "kind", "writeAuthority");
        if (!KINDS.contains(kind)) throw ServiceException.validation("writeAuthority.kind is unsupported: " + kind);
        requireDirectCandidate(intent, sheetId, params, value);
        validateDirectEntry(root, sheet, sheetId, params, value, intent);
    }

    private static ObjectNode requireIntent(ObjectNode params) {
        JsonNode value = params.get("writeAuthority");
        if (value == null || !value.isObject()) throw ServiceException.validation("writeAuthority is required for cell writes");
        ObjectNode intent = (ObjectNode) value;
        JsonNode decision = intent.get("validationDecision");
        if (decision == null || !decision.isObject()) throw ServiceException.validation("writeAuthority.validationDecision is required");
        String status = requiredText((ObjectNode) decision, "status", "writeAuthority.validationDecision");
        if (!List.of("accepted", "confirmed").contains(status)) {
            throw ServiceException.validation("writeAuthority.validationDecision.status is unsupported: " + status);
        }
        return intent;
    }

    private static void requireDirectCandidate(ObjectNode intent, String sheetId, ObjectNode params, ObjectNode value) {
        JsonNode target = intent.get("target");
        if (target == null || !target.isObject()
                || !sheetId.equals(target.path("sheetId").asText())
                || target.path("row").asInt(-1) != params.path("row").asInt(-2)
                || target.path("column").asInt(-1) != params.path("column").asInt(-2)) {
            throw ServiceException.validation("writeAuthority.target does not match cell.set target");
        }
        JsonNode candidate = intent.get("candidate");
        if (candidate == null || !candidate.isObject() || !candidate.equals(value)) {
            throw ServiceException.validation("writeAuthority.candidate does not match cell.set value");
        }
    }

    private static void validateDirectEntry(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params, ObjectNode value, ObjectNode intent) {
        int row = params.path("row").asInt(-1);
        int column = params.path("column").asInt(-1);
        ObjectNode rule = findRule(root, sheet, sheetId, row, column);
        if (rule == null) {
            requireDecision(intent, true, null, null);
            return;
        }
        JsonNode formula = value.get("formula");
        if (formula != null && !formula.isNull()) {
            throw ServiceException.unsupportedFeature("UNSUPPORTED_FEATURE: formula entry under data validation requires the shared calculation evaluator");
        }
        String ruleId = textOrNull(rule, "id");
        JsonNode declaredRuleId = intent.path("validationDecision").get("ruleId");
        if (declaredRuleId != null && !declaredRuleId.isNull() && !declaredRuleId.isTextual()) {
            throw ServiceException.validation("writeAuthority.validationDecision.ruleId must be a string");
        }
        if (declaredRuleId != null && !declaredRuleId.isNull() && !declaredRuleId.asText().equals(ruleId)) {
            throw ServiceException.validation("CELL_ENTRY_VALIDATION_STALE: validation rule changed before commit");
        }
        ValidationResult result = evaluateRule(root, sheet, sheetId, row, column, value, rule);
        requireDecision(intent, result.valid(), result.alertStyle(), result.message());
    }

    private static ObjectNode findRule(ObjectNode root, ObjectNode sheet, String sheetId, int row, int column) {
        JsonNode rules = sheet.get("dataValidations");
        if (rules == null) return null;
        if (!rules.isArray()) throw ServiceException.validation("dataValidations must be an array");
        ObjectNode matched = null;
        for (JsonNode ruleNode : rules) {
            if (!ruleNode.isObject()) throw ServiceException.validation("dataValidations contains a non-object rule");
            JsonNode ranges = ruleNode.get("ranges");
            if (ranges == null || !ranges.isArray()) throw ServiceException.validation("data validation ranges are required");
            for (JsonNode rangeNode : ranges) {
                RangeRef range = SnapshotMutationSupport.range(root, rangeNode);
                if (sheetId.equals(range.sheetId()) && row >= range.startRow() && row <= range.endRow()
                        && column >= range.startColumn() && column <= range.endColumn()) {
                    if (matched != null) throw ServiceException.validation("Multiple data validation rules apply to one cell");
                    matched = (ObjectNode) ruleNode;
                }
            }
        }
        return matched;
    }

    private static ValidationResult evaluateRule(ObjectNode root, ObjectNode sheet, String sheetId, int row, int column, ObjectNode value, ObjectNode rule) {
        String type = requiredText(rule, "type", "data validation rule").toLowerCase(Locale.ROOT);
        String alertStyle = rule.path("alertStyle").asText("stop").toLowerCase(Locale.ROOT);
        if (!List.of("stop", "warning", "information").contains(alertStyle)) throw ServiceException.validation("Unsupported data validation alertStyle: " + alertStyle);
        JsonNode candidate = value.get("value");
        boolean blank = candidate == null || candidate.isNull() || (candidate.isTextual() && candidate.asText().isEmpty());
        if (blank) return new ValidationResult(rule.path("allowBlank").asBoolean(true), alertStyle, "该单元格不允许为空");
        if ("list".equals(type)) {
            if (!(candidate.isTextual() || candidate.isNumber() || candidate.isBoolean())) {
                throw ServiceException.unsupportedFeature("UNSUPPORTED_FEATURE: data validation cannot evaluate a non-scalar cell value");
            }
            boolean valid = listContains(root, rule, candidate.asText());
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
        if ("custom".equals(type)) throw ServiceException.unsupportedFeature("UNSUPPORTED_FEATURE: server cannot authoritatively evaluate custom data validation formulas");
        throw ServiceException.unsupportedFeature("UNSUPPORTED_FEATURE: server cannot authoritatively evaluate data validation type " + type);
    }

    private static boolean listContains(ObjectNode root, ObjectNode rule, String candidate) {
        JsonNode source = rule.get("listSource");
        if (source != null && source.isObject()) {
            String kind = source.path("kind").asText();
            if ("values".equals(kind)) {
                JsonNode values = source.get("values");
                if (values == null || !values.isArray()) throw ServiceException.validation("list validation values are required");
                for (JsonNode value : values) {
                    if (!value.isTextual()) throw ServiceException.validation("list validation values must be strings");
                    if (value.asText().equalsIgnoreCase(candidate)) return true;
                }
                return false;
            }
            if ("range".equals(kind)) {
                RangeRef range = SnapshotMutationSupport.range(root, source.get("range"));
                ObjectNode sourceSheet = SnapshotMutationSupport.sheet(root, range.sheetId());
                JsonNode cells = sourceSheet.get("cells");
                if (cells == null || !cells.isObject()) throw ServiceException.validation("Validation list source cells are invalid");
                Iterator<Map.Entry<String, JsonNode>> rows = cells.fields();
                while (rows.hasNext()) {
                    Map.Entry<String, JsonNode> rowEntry = rows.next();
                    int row = storedCoordinate(rowEntry.getKey(), SnapshotMutationSupport.MAX_ROW, "row");
                    if (row < range.startRow() || row > range.endRow()) continue;
                    if (!rowEntry.getValue().isObject()) throw ServiceException.validation("Validation list source row is invalid");
                    Iterator<Map.Entry<String, JsonNode>> columns = rowEntry.getValue().fields();
                    while (columns.hasNext()) {
                        Map.Entry<String, JsonNode> columnEntry = columns.next();
                        int column = storedCoordinate(columnEntry.getKey(), SnapshotMutationSupport.MAX_COLUMN, "column");
                        if (column < range.startColumn() || column > range.endColumn()) continue;
                        JsonNode cell = columnEntry.getValue();
                        if (!cell.isObject()) throw ServiceException.validation("Validation list source cell is invalid");
                        JsonNode value = cell.has("formulaValue") ? cell.get("formulaValue") : cell.get("value");
                        if (value == null || value.isNull() || isFormulaError(value)) continue;
                        if (!(value.isTextual() || value.isNumber() || value.isBoolean())) {
                            throw ServiceException.unsupportedFeature("UNSUPPORTED_FEATURE: validation list source contains a non-scalar value");
                        }
                        if (value.asText().equalsIgnoreCase(candidate)) return true;
                    }
                }
                return false;
            }
            throw ServiceException.unsupportedFeature("UNSUPPORTED_FEATURE: formula-backed list validation requires shared evaluator authority");
        }
        String formula = textOrNull(rule, "formula1");
        if (formula == null) throw ServiceException.validation("list validation source is required");
        for (String item : formula.replaceFirst("^=", "").split(",", -1)) {
            if (!item.isBlank() && item.trim().replace("\"", "").equalsIgnoreCase(candidate)) return true;
        }
        return false;
    }

    private static int storedCoordinate(String value, int maximum, String label) {
        try {
            int coordinate = Integer.parseInt(value);
            if (coordinate < 0 || coordinate > maximum || !Integer.toString(coordinate).equals(value)) {
                throw ServiceException.validation("Validation list source " + label + " key is invalid");
            }
            return coordinate;
        } catch (NumberFormatException exception) {
            throw ServiceException.validation("Validation list source " + label + " key is invalid");
        }
    }

    private static boolean isFormulaError(JsonNode value) {
        return value.isObject() && "error".equals(value.path("kind").asText());
    }

    private static ValidationResult compareNumeric(double actual, ObjectNode rule, String alertStyle) {
        String operator = rule.path("operator").asText("between");
        Double first = numberOrNull(rule.get("formula1"), "formula1");
        Double second = numberOrNull(rule.get("formula2"), "formula2");
        boolean valid = switch (operator) {
            case "greaterThan" -> actual > requiredBound(first, "formula1");
            case "lessThan" -> actual < requiredBound(first, "formula1");
            case "equal" -> actual == requiredBound(first, "formula1");
            case "notEqual" -> actual != requiredBound(first, "formula1");
            case "notBetween" -> actual < requiredBound(first, "formula1") || actual > requiredBound(second, "formula2");
            case "between" -> actual >= requiredBound(first, "formula1") && actual <= requiredBound(second, "formula2");
            default -> throw ServiceException.validation("Unsupported data validation operator: " + operator);
        };
        return new ValidationResult(valid, alertStyle, "不符合数据验证规则");
    }

    private static double requiredBound(Double value, String field) {
        if (value == null) throw ServiceException.validation("Data validation " + field + " is required for its operator");
        return value;
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

    private static Double numberOrNull(JsonNode value, String field) {
        if (value == null || value.isNull()) return null;
        if (value.isNumber()) {
            double number = value.asDouble();
            if (Double.isFinite(number)) return number;
            throw ServiceException.validation("Data validation " + field + " must be finite");
        }
        if (!value.isTextual()) throw ServiceException.validation("Data validation " + field + " must be a formula or numeric constant");
        String text = value.asText().trim().replaceFirst("^=", "");
        try {
            double number = Double.parseDouble(text);
            if (Double.isFinite(number)) return number;
            throw ServiceException.validation("Data validation " + field + " must be finite");
        } catch (NumberFormatException exception) {
            throw ServiceException.unsupportedFeature("UNSUPPORTED_FEATURE: formula-based data validation bounds require the shared calculation evaluator");
        }
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
