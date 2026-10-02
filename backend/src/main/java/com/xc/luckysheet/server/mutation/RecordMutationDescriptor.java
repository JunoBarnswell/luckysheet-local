package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.*;
import com.xc.luckysheet.server.service.ServiceException;
import java.util.*;

final class RecordMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of("table.configure", "relationship.set", "relationship.remove", "record.set", "record.restore");
    RecordMutationDescriptor(String id) { super(id, WorkbookAclRole.EDITOR); }
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        if (id().equals("record.set") || id().equals("record.restore")) {
            SnapshotMutationSupport.validateKnownKeys(params, id().equals("record.set") ? Set.of("tableId", "recordId", "fieldId", "value", "writeAuthority") : Set.of("tableId", "recordId", "fieldId", "previous"), id());
            JsonNode table = RecordTableValidator.table(snapshot, SnapshotMutationSupport.text(params, "tableId")), range = table.path("sourceRange");
            JsonNode field = RecordTableValidator.field(table, SnapshotMutationSupport.text(params, "fieldId"));
            Integer row = RecordTableValidator.rows(snapshot, table).get(SnapshotMutationSupport.text(params, "recordId"));
            if (row == null) throw ServiceException.notFound("Record not found");
            if (field.has("calculation") || field.path("id").asText().equals(table.path("recordIdFieldId").asText())) throw ServiceException.unsupportedFeature("Computed and identity fields are read-only");
            JsonNode previous = params.get("previous");
            if (id().equals("record.restore") && (previous == null || !(previous.isNull() || previous.isObject()) || previous.has("formula"))) throw ServiceException.validation("Record restore requires a stored cell or null");
            JsonNode value = id().equals("record.set") ? params.get("value") : previous.isNull() || !previous.has("value") ? com.fasterxml.jackson.databind.node.NullNode.instance : previous.get("value");
            if (value == null || !(value.isNull() || value.isTextual() || value.isNumber() || value.isBoolean())) throw ServiceException.validation("Record field value must be scalar");
            int column = range.path("startColumn").asInt() + field.path("ordinal").asInt();
            String sheetId = range.path("sheetId").asText();
            if (!mutation.sheetId().equals(sheetId)) throw ServiceException.validation("Record source identity mismatch");
            ObjectNode source = (ObjectNode) RecordTableValidator.sheet(snapshot, sheetId);
            JsonNode existing = source.path("cells").path(Integer.toString(row)).path(Integer.toString(column));
            ObjectNode candidate = existing.isObject() ? ((ObjectNode) existing).deepCopy() : source.objectNode();
            candidate.set("value", value.deepCopy()); candidate.remove("formula"); candidate.remove("formulaValue");
            ObjectNode authorityParams = source.objectNode();
            authorityParams.put("sheetId", sheetId); authorityParams.put("row", row); authorityParams.put("column", column);
            authorityParams.set("writeAuthority", params.get("writeAuthority"));
            if (id().equals("record.set")) CellWriteAuthority.requireCellWrite(SnapshotMutationSupport.root(snapshot), source, sheetId, authorityParams, candidate);
            return List.of(new RangeRef(sheetId, row, row, column, column));
        }
        if (id().equals("table.configure")) {
            SnapshotMutationSupport.validateKnownKeys(params, Set.of("table"), id());
            JsonNode table = SnapshotMutationSupport.requiredObject(params, "table");
            JsonNode previous = RecordTableValidator.table(snapshot, table.path("id").asText());
            if (!previous.path("sourceRange").equals(table.path("sourceRange")) || !previous.path("sourceId").equals(table.path("sourceId"))) throw ServiceException.unsupportedFeature("Record table configuration cannot move source ownership");
            if (!previous.path("sourceSheetId").equals(table.path("sourceSheetId")) || !previous.path("rowCount").equals(table.path("rowCount")) || previous.path("fields").size() != table.path("fields").size()) throw ServiceException.conflict("RECORD_FIELD_IDENTITY_IMMUTABLE");
            for (JsonNode old : previous.path("fields")) if (!old.path("ordinal").equals(RecordTableValidator.field(table, old.path("id").asText()).path("ordinal"))) throw ServiceException.conflict("RECORD_FIELD_IDENTITY_IMMUTABLE");
            RecordTableValidator.validate(snapshot, table);
            return List.of(SnapshotMutationSupport.range(SnapshotMutationSupport.root(snapshot), table.get("sourceRange")));
        }
        if (id().equals("relationship.set")) {
            SnapshotMutationSupport.validateKnownKeys(params, Set.of("relationship"), id());
            JsonNode relation = SnapshotMutationSupport.requiredObject(params, "relationship");
            RecordTableValidator.validateRelationship(snapshot, relation);
            return relationRanges(snapshot, relation);
        }
        SnapshotMutationSupport.validateKnownKeys(params, Set.of("relationshipId"), id());
        JsonNode relation = SnapshotMutationSupport.requireById(SnapshotMutationSupport.dataModelArray(SnapshotMutationSupport.root(snapshot), "relationships"), SnapshotMutationSupport.text(params, "relationshipId"), "Relationship");
        for (JsonNode table : snapshot.path("dataModel").path("tables")) for (JsonNode field : table.path("fields")) if (field.path("calculation").path("relationshipId").asText().equals(params.path("relationshipId").asText())) throw ServiceException.conflict("Relationship is owned by a calculated field");
        return relationRanges(snapshot, relation);
    }
    private List<RangeRef> relationRanges(JsonNode snapshot, JsonNode relation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        return List.of(SnapshotMutationSupport.range(root, RecordTableValidator.table(snapshot, relation.path("fromTableId").asText()).get("sourceRange")), SnapshotMutationSupport.range(root, RecordTableValidator.table(snapshot, relation.path("toTableId").asText()).get("sourceRange")));
    }
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        var ranges = affectedRanges(snapshot, mutation);
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy()), params = SnapshotMutationSupport.params(mutation);
        if (id().equals("record.set") || id().equals("record.restore")) {
            RangeRef range = ranges.getFirst(); ObjectNode cell = SnapshotMutationSupport.object(SnapshotMutationSupport.object(SnapshotMutationSupport.object((ObjectNode) RecordTableValidator.sheet(root, range.sheetId()), "cells"), Integer.toString(range.startRow())), Integer.toString(range.startColumn()));
            if (id().equals("record.set")) { cell.set("value", params.get("value").deepCopy()); cell.remove("formula"); cell.remove("formulaValue"); }
            else {
                ObjectNode row;
                ObjectNode cells = (ObjectNode) RecordTableValidator.sheet(root, range.sheetId()).path("cells");
                row = (ObjectNode) cells.get(Integer.toString(range.startRow()));
                if (params.get("previous").isNull()) { row.remove(Integer.toString(range.startColumn())); if (row.isEmpty()) cells.remove(Integer.toString(range.startRow())); }
                else row.set(Integer.toString(range.startColumn()), params.get("previous").deepCopy());
            }
        } else if (id().equals("table.configure") || id().equals("relationship.set")) {
            String collection = id().equals("table.configure") ? "tables" : "relationships", key = id().equals("table.configure") ? "table" : "relationship";
            var items = SnapshotMutationSupport.dataModelArray(root, collection); JsonNode value = params.get(key);
            int index = SnapshotMutationSupport.indexById(items, value.path("id").asText());
            if (index < 0) items.add(value.deepCopy()); else items.set(index, value.deepCopy());
        } else {
            var items = SnapshotMutationSupport.dataModelArray(root, "relationships"); items.remove(SnapshotMutationSupport.indexById(items, params.path("relationshipId").asText()));
        }
        RecordTableValidator.validateWorkbook(root);
        return root;
    }
}
