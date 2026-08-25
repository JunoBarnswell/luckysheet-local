package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Resolves every canonical Pivot source against the candidate workbook
 * snapshot. A Pivot source is not valid merely because its identifier has the
 * right JSON type: the referenced object and its concrete source range must
 * exist before a mutation can be committed.
 */
final class PivotSourceResolver {
    private static final Pattern CELL = Pattern.compile("\\$?([A-Za-z]{1,3})\\$?([0-9]+)");

    private PivotSourceResolver() {
    }

    record Resolution(List<RangeRef> ranges, List<String> fieldIds) {
        Resolution {
            ranges = List.copyOf(ranges);
            fieldIds = List.copyOf(fieldIds);
        }
    }

    static Resolution resolve(ObjectNode root, ObjectNode source) {
        String kind = SnapshotMutationSupport.text(source, "kind");
        return switch (kind) {
            case "worksheet-range" -> {
                RangeRef range = SnapshotMutationSupport.range(root, source.get("range"));
                List<String> fieldIds = new ArrayList<>();
                for (int ordinal = 0; ordinal <= range.endColumn() - range.startColumn(); ordinal++) {
                    fieldIds.add("sheet:" + range.sheetId() + ":column:" + (range.startColumn() + ordinal) + ":range:0");
                }
                yield new Resolution(List.of(range), fieldIds);
            }
            case "worksheet-ranges" -> resolveWorksheetRanges(root, source);
            case "table" -> resolveTable(root, SnapshotMutationSupport.text(source, "tableId"));
            case "named-range" -> resolveNamedRange(root, SnapshotMutationSupport.text(source, "name"),
                    source.has("sheetId") ? SnapshotMutationSupport.text(source, "sheetId") : null);
            case "data-source" -> resolveDataSource(root, SnapshotMutationSupport.text(source, "dataSourceId"));
            default -> throw ServiceException.validation("Pivot source kind is invalid");
        };
    }

    private static Resolution resolveWorksheetRanges(ObjectNode root, ObjectNode source) {
        ArrayNode ranges = SnapshotMutationSupport.requiredArray(source, "ranges");
        List<RangeRef> resolved = new ArrayList<>();
        List<String> fieldIds = new ArrayList<>();
        Set<String> sourceIds = new HashSet<>();
        for (JsonNode raw : ranges) {
            if (!raw.isObject()) throw ServiceException.validation("Pivot source range must be an object");
            ObjectNode entry = (ObjectNode) raw;
            String sourceId = SnapshotMutationSupport.text(entry, "sourceId");
            if (!sourceIds.add(sourceId)) throw ServiceException.validation("Pivot sourceId is duplicated: " + sourceId);
            RangeRef range = SnapshotMutationSupport.range(root, entry.get("range"));
            resolved.add(range);
            for (int ordinal = 0; ordinal <= range.endColumn() - range.startColumn(); ordinal++) {
                fieldIds.add("source:" + sourceId + ":column:" + ordinal);
            }
        }
        if (resolved.isEmpty()) throw ServiceException.validation("Pivot source ranges are empty");
        return new Resolution(resolved, fieldIds);
    }

    private static Resolution resolveTable(ObjectNode root, String tableId) {
        JsonNode dataModel = root.get("dataModel");
        if (dataModel != null && dataModel.isObject()) {
            JsonNode tables = dataModel.get("tables");
            if (tables != null && tables.isArray()) {
                for (JsonNode raw : tables) {
                    if (!raw.isObject() || !tableId.equals(raw.path("id").asText())) continue;
                    JsonNode sourceRange = raw.get("sourceRange");
                    if (sourceRange != null && !sourceRange.isNull()) {
                        RangeRef range = SnapshotMutationSupport.range(root, sourceRange);
                        return fromRanges(List.of(range), "table:" + tableId + ":column:");
                    }
                    // A workbook table without a sheet-backed range is not a
                    // resolvable local Pivot source. The sheet-table lookup
                    // below still permits the canonical fallback used by the
                    // frontend model.
                    break;
                }
            }
        }
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject()) continue;
            JsonNode sheetTables = rawSheet.get("sheetTables");
            if (sheetTables == null || !sheetTables.isArray()) continue;
            for (JsonNode raw : sheetTables) {
                if (!raw.isObject()) continue;
                if (!tableId.equals(raw.path("id").asText()) && !tableId.equals(raw.path("name").asText())) continue;
                RangeRef range = SnapshotMutationSupport.range(root, raw.get("range"));
                return fromRanges(List.of(range), "table:" + tableId + ":column:");
            }
        }
        throw ServiceException.notFound("Pivot table source not found: " + tableId);
    }

    private static Resolution resolveDataSource(ObjectNode root, String sourceId) {
        JsonNode dataModel = root.get("dataModel");
        if (dataModel == null || !dataModel.isObject()) {
            throw ServiceException.notFound("Pivot data source not found: " + sourceId);
        }
        JsonNode rawSources = dataModel.get("sources");
        if (rawSources == null || !rawSources.isArray()) {
            throw ServiceException.notFound("Pivot data source not found: " + sourceId);
        }
        ArrayNode sources = (ArrayNode) rawSources;
        for (JsonNode raw : sources) {
            if (!raw.isObject() || !sourceId.equals(raw.path("id").asText())) continue;
            ObjectNode source = (ObjectNode) raw;
            List<RangeRef> ranges = new ArrayList<>();
            JsonNode single = source.get("sourceRange");
            if (single != null && !single.isNull()) ranges.add(SnapshotMutationSupport.range(root, single));
            JsonNode many = source.get("ranges");
            if (many != null) {
                if (!many.isArray()) throw ServiceException.validation("Pivot data source ranges must be an array");
                for (JsonNode range : many) ranges.add(SnapshotMutationSupport.range(root, range));
            }
            List<String> fieldIds = new ArrayList<>();
            JsonNode fields = source.get("fields");
            if (fields != null) {
                if (!fields.isArray()) throw ServiceException.validation("Pivot data source fields must be an array");
                for (JsonNode field : fields) {
                    if (!field.isObject()) throw ServiceException.validation("Pivot data source field must be an object");
                    String fieldId = SnapshotMutationSupport.text((ObjectNode) field, "id");
                    if (!fieldIds.contains(fieldId)) fieldIds.add(fieldId);
                }
            }
            return new Resolution(ranges, fieldIds);
        }
        throw ServiceException.notFound("Pivot data source not found: " + sourceId);
    }

    private static Resolution resolveNamedRange(ObjectNode root, String name, String scopeSheetId) {
        if (scopeSheetId != null) SnapshotMutationSupport.sheet(root, scopeSheetId);
        String normalizedName = name.trim();
        ObjectNode model = findDefinedName(root, normalizedName, scopeSheetId);
        String formula = SnapshotMutationSupport.text(model, "formula");
        String fallbackSheet = scopeSheetId != null ? scopeSheetId : firstSheetId(root);
        RangeRef range = parseA1Range(root, formula, fallbackSheet);
        String identity = "name:" + (scopeSheetId == null ? "*" : scopeSheetId) + ":" + normalizedName + ":column:";
        return fromRanges(List.of(range), identity);
    }

    private static ObjectNode findDefinedName(ObjectNode root, String name, String scopeSheetId) {
        JsonNode rawModels = root.get("definedNameModels");
        if (rawModels == null || !rawModels.isArray()) {
            throw ServiceException.notFound("Pivot named range not found: " + name);
        }
        for (JsonNode raw : rawModels) {
            if (!raw.isObject() || !name.equalsIgnoreCase(raw.path("name").asText())) continue;
            String scope = raw.path("scope").asText();
            boolean matches = scopeSheetId == null
                    ? "workbook".equals(scope) && !raw.has("sheetId")
                    : "sheet".equals(scope) && scopeSheetId.equals(raw.path("sheetId").asText());
            if (matches) return (ObjectNode) raw;
        }
        throw ServiceException.notFound("Pivot named range not found: " + name
                + (scopeSheetId == null ? "" : " on sheet " + scopeSheetId));
    }

    private static RangeRef parseA1Range(ObjectNode root, String formula, String fallbackSheetId) {
        String expression = formula.trim();
        if (expression.startsWith("=")) expression = expression.substring(1).trim();
        if (expression.isBlank() || expression.endsWith("#") || expression.contains(",") || expression.contains(";")) {
            throw ServiceException.validation("Pivot named range formula is not a single resolvable range");
        }
        String sheetToken = null;
        int bang = findSheetSeparator(expression);
        String address = expression;
        if (bang >= 0) {
            sheetToken = expression.substring(0, bang);
            address = expression.substring(bang + 1);
        }
        String sheetId = sheetToken == null ? fallbackSheetId : resolveSheetToken(root, sheetToken);
        if (sheetId == null || sheetId.isBlank()) throw ServiceException.validation("Pivot named range sheet is invalid");
        Matcher matcher = CELL.matcher(address);
        if (!matcher.find() || matcher.start() != 0) throw ServiceException.validation("Pivot named range formula is invalid");
        int colon = address.indexOf(':');
        String first = colon < 0 ? address : address.substring(0, colon);
        String second = colon < 0 ? first : address.substring(colon + 1);
        Matcher start = CELL.matcher(first);
        Matcher end = CELL.matcher(second);
        if (!start.matches() || !end.matches()) throw ServiceException.validation("Pivot named range formula is invalid");
        int startColumn = columnIndex(start.group(1));
        int endColumn = columnIndex(end.group(1));
        int startRow = rowIndex(start.group(2));
        int endRow = rowIndex(end.group(2));
        if (endRow < startRow || endColumn < startColumn) throw ServiceException.validation("Pivot named range formula has reversed bounds");
        ObjectNode range = JsonNodeFactory.instance.objectNode()
                .put("sheetId", sheetId)
                .put("startRow", startRow)
                .put("endRow", endRow)
                .put("startColumn", startColumn)
                .put("endColumn", endColumn);
        return SnapshotMutationSupport.range(root, range);
    }

    private static int findSheetSeparator(String expression) {
        boolean quoted = false;
        for (int index = 0; index < expression.length(); index++) {
            char character = expression.charAt(index);
            if (character == '\'') quoted = !quoted;
            else if (character == '!' && !quoted) return index;
        }
        return -1;
    }

    private static String resolveSheetToken(ObjectNode root, String token) {
        String normalized = token.trim();
        if (normalized.startsWith("'") && normalized.endsWith("'") && normalized.length() >= 2) {
            normalized = normalized.substring(1, normalized.length() - 1).replace("''", "'");
        }
        for (JsonNode sheet : SnapshotMutationSupport.sheets(root)) {
            if (!sheet.isObject()) continue;
            if (normalized.equals(sheet.path("id").asText()) || normalized.equalsIgnoreCase(sheet.path("name").asText())) {
                return sheet.path("id").asText();
            }
        }
        throw ServiceException.notFound("Pivot named range sheet not found: " + normalized);
    }

    private static String firstSheetId(ObjectNode root) {
        JsonNode first = SnapshotMutationSupport.sheets(root).get(0);
        if (first == null || !first.isObject() || first.path("id").asText().isBlank()) {
            throw ServiceException.validation("Workbook has no primary sheet");
        }
        return first.path("id").asText();
    }

    private static int columnIndex(String value) {
        int result = 0;
        for (int index = 0; index < value.length(); index++) result = result * 26 + (Character.toUpperCase(value.charAt(index)) - 'A' + 1);
        result--;
        if (result < 0 || result > SnapshotMutationSupport.MAX_COLUMN) throw ServiceException.validation("Pivot named range column is out of bounds");
        return result;
    }

    private static int rowIndex(String value) {
        try {
            int result = Integer.parseInt(value) - 1;
            if (result < 0 || result > SnapshotMutationSupport.MAX_ROW) throw ServiceException.validation("Pivot named range row is out of bounds");
            return result;
        } catch (NumberFormatException error) {
            throw ServiceException.validation("Pivot named range row is invalid");
        }
    }

    private static Resolution fromRanges(List<RangeRef> ranges, String identityPrefix) {
        List<String> fieldIds = new ArrayList<>();
        for (RangeRef range : ranges) {
            int width = range.endColumn() - range.startColumn() + 1;
            for (int ordinal = 0; ordinal < width; ordinal++) {
                fieldIds.add(identityPrefix + ordinal);
            }
        }
        return new Resolution(ranges, fieldIds);
    }

}
