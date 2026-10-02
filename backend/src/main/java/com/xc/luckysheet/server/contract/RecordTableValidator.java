package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.service.ServiceException;
import java.util.*;

/** Record identities and field ownership are checked at every canonical persistence boundary. */
public final class RecordTableValidator {
    private RecordTableValidator() {}
    public static JsonNode table(JsonNode root, String id) {
        for (JsonNode table : root.path("dataModel").path("tables")) if (table.path("id").asText().equals(id)) return table;
        throw ServiceException.notFound("Record table not found");
    }
    public static JsonNode sheet(JsonNode root, String id) {
        for (JsonNode sheet : root.path("sheets")) if (sheet.path("id").asText().equals(id)) return sheet;
        throw ServiceException.notFound("Record source not found");
    }
    public static JsonNode field(JsonNode table, String id) {
        for (JsonNode field : table.path("fields")) if (field.path("id").asText().equals(id)) return field;
        throw ServiceException.notFound("Record field not found");
    }
    public static Map<String, Integer> rows(JsonNode root, JsonNode table) {
        String identityId = table.path("recordIdFieldId").asText();
        if (identityId.isBlank() || !table.path("sourceRange").isObject()) throw ServiceException.validation("Record table needs an identity field and worksheet source");
        JsonNode identity = field(table, identityId), range = table.path("sourceRange");
        if (!identity.path("type").asText().equals("text") || identity.has("calculation")) throw ServiceException.validation("Record identity must be a stored text field");
        int start = range.path("startRow").asInt(), end = range.path("endRow").asInt();
        if (end - start > 100000) throw ServiceException.unsupportedFeature("Record table exceeds bounded worksheet source contract");
        int column = range.path("startColumn").asInt() + identity.path("ordinal").asInt();
        JsonNode source = sheet(root, range.path("sheetId").asText());
        Map<String, Integer> rows = new LinkedHashMap<>();
        for (int row = start + 1; row <= end; row++) {
            JsonNode id = source.path("cells").path(Integer.toString(row)).path(Integer.toString(column)).get("value");
            if (id == null || !id.isTextual() || id.asText().isBlank() || id.asText().length() > 200 || rows.putIfAbsent(id.asText(), row) != null) throw ServiceException.validation("Record identity is missing or duplicated");
        }
        return rows;
    }
    public static void validate(JsonNode root, JsonNode table) {
        boolean computed = false;
        for (JsonNode field : table.path("fields")) if (field.has("calculation")) computed = true;
        if (!computed && !table.has("recordIdFieldId")) return;
        Map<String, Integer> rows = rows(root, table);
        JsonNode range = table.path("sourceRange"), source = sheet(root, range.path("sheetId").asText());
        Set<String> ids = new HashSet<>(); Set<Integer> ordinals = new HashSet<>();
        for (JsonNode field : table.path("fields")) {
            String id = field.path("id").asText(); int ordinal = field.path("ordinal").asInt(-1);
            if (id.isBlank() || !ids.add(id) || !field.path("ordinal").isIntegralNumber() || ordinal < 0 || !ordinals.add(ordinal)
                || range.path("startColumn").asInt() + ordinal > range.path("endColumn").asInt()) throw ServiceException.validation("Record field identities/ordinals are invalid");
            JsonNode calculation = field.get("calculation");
            if (calculation == null) {
                for (int row : rows.values()) if (source.path("cells").path(Integer.toString(row)).path(Integer.toString(range.path("startColumn").asInt() + ordinal)).has("formula")) throw ServiceException.validation("Stored Record fields cannot own cell formulas");
                continue;
            }
            String kind = calculation.path("kind").asText();
            Set<String> keys = kind.equals("formula") ? Set.of("kind", "formula") : kind.equals("lookup") ? Set.of("kind", "relationshipId", "targetFieldId", "direction") : kind.equals("rollup") ? Set.of("kind", "relationshipId", "targetFieldId", "direction", "aggregate") : Set.of();
            if (!calculation.isObject() || keys.isEmpty() || calculation.size() != keys.size()) throw ServiceException.validation("Record calculation is invalid");
            calculation.fieldNames().forEachRemaining(key -> { if (!keys.contains(key)) throw ServiceException.validation("Record calculation has unsupported fields"); });
            if (kind.equals("formula")) {
                String formula = calculation.path("formula").asText();
                if (!formula.startsWith("=") || formula.length() > 32768) throw ServiceException.validation("Record formula is invalid");
            } else {
                if (calculation.path("relationshipId").asText().isBlank() || calculation.path("targetFieldId").asText().isBlank()
                    || !Set.of("forward", "reverse").contains(calculation.path("direction").asText())
                    || kind.equals("rollup") && !Set.of("SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "PRODUCT").contains(calculation.path("aggregate").asText())) throw ServiceException.validation("Record relation calculation is invalid");
            }
            int column = range.path("startColumn").asInt() + ordinal;
            for (int row : rows.values()) {
                JsonNode cell = source.path("cells").path(Integer.toString(row)).path(Integer.toString(column));
                JsonNode value = cell.get("value");
                if (cell.has("formula") || value != null && !value.isNull() && !(value.isTextual() && value.asText().isEmpty())) throw ServiceException.conflict("Computed fields cannot overlap stored inputs");
            }
        }
        for (int ordinal = 0; ordinal < ids.size(); ordinal++) if (!ordinals.contains(ordinal)) throw ServiceException.validation("Record field ordinals must be contiguous");
    }
    public static void validateCalculationOwners(JsonNode root) {
        for (JsonNode table : root.path("dataModel").path("tables")) for (JsonNode field : table.path("fields")) {
            JsonNode calculation = field.path("calculation");
            if (calculation.isMissingNode() || calculation.path("kind").asText().equals("formula")) continue;
            JsonNode relation = null;
            for (JsonNode candidate : root.path("dataModel").path("relationships")) if (candidate.path("id").asText().equals(calculation.path("relationshipId").asText())) relation = candidate;
            boolean forward = calculation.path("direction").asText().equals("forward");
            if (relation == null || !table.path("id").equals(relation.path(forward ? "fromTableId" : "toTableId"))) throw ServiceException.validation("Calculated field relationship owner is invalid");
            JsonNode target = table(root, relation.path(forward ? "toTableId" : "fromTableId").asText());
            if (!target.has("recordIdFieldId")) throw ServiceException.validation("Calculated field target is not a Record table");
            field(target, calculation.path("targetFieldId").asText());
        }
    }
    public static void validateRelationship(JsonNode root, JsonNode relation) {
        JsonNode from = table(root, relation.path("fromTableId").asText()), to = table(root, relation.path("toTableId").asText());
        JsonNode field = field(from, relation.path("fromFieldId").asText());
        if (relation.path("id").asText().isBlank() || relation.size() != 6 || field.has("calculation")
            || !relation.path("toFieldId").asText().equals(to.path("recordIdFieldId").asText())
            || !Set.of("many-to-one", "one-to-one").contains(relation.path("cardinality").asText())) throw ServiceException.validation("Relation must reference immutable target record IDs");
        Map<String, Integer> targets = rows(root, to); Set<String> used = new HashSet<>();
        JsonNode range = from.path("sourceRange"), source = sheet(root, range.path("sheetId").asText());
        int column = range.path("startColumn").asInt() + field.path("ordinal").asInt();
        for (int row : rows(root, from).values()) {
            JsonNode id = source.path("cells").path(Integer.toString(row)).path(Integer.toString(column)).get("value");
            if (id == null || id.isNull() || id.isTextual() && id.asText().isEmpty()) continue;
            if (!id.isTextual() || !targets.containsKey(id.asText()) || relation.path("cardinality").asText().equals("one-to-one") && !used.add(id.asText())) throw ServiceException.validation("Relation has a missing or duplicate target record");
        }
    }
    public static void validateWorkbook(JsonNode root) {
        for (JsonNode table : root.path("dataModel").path("tables")) validate(root, table);
        validateCalculationOwners(root);
        for (JsonNode relation : root.path("dataModel").path("relationships")) {
            // Ordinary relational data models keep their existing join-key contract.
            if (table(root, relation.path("fromTableId").asText()).has("recordIdFieldId")) validateRelationship(root, relation);
        }
    }
    public static void guardWrites(JsonNode root, String mutation, List<RangeRef> ranges) {
        if (!Set.of("cell.set", "cell.restore", "range.set", "range.paste", "range.clear", "range.clear.restore", "fill.applied", "fill.restored", "query.load.range", "range.move", "record.set").contains(mutation)) return;
        if (mutation.equals("record.set")) return;
        for (JsonNode table : root.path("dataModel").path("tables")) {
            if (!table.has("recordIdFieldId")) continue;
            JsonNode source = table.path("sourceRange");
            for (JsonNode field : table.path("fields")) {

                int column = source.path("startColumn").asInt() + field.path("ordinal").asInt();
                for (RangeRef range : ranges) if (range.sheetId().equals(source.path("sheetId").asText()) && range.startRow() <= source.path("endRow").asInt() && range.endRow() > source.path("startRow").asInt() && range.startColumn() <= column && range.endColumn() >= column) throw ServiceException.unsupportedFeature("Record fields require writes addressed by Record/Field IDs");
            }
        }
    }
}
