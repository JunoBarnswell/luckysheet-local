package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.regex.Pattern;

/** Reducers for persisted query definitions and range-backed query loads. */
final class QueryMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of("query.definition.replace", "query.load.range", "query.load.sheet-table", "query.load.pivot-source");
    private static final Set<String> STEP_KINDS = Set.of("source", "filter", "select-columns", "rename-column", "sort", "group-by", "join", "pivot", "custom");
    private static final Pattern SECRET_KEY = Pattern.compile("(?:pass(word)?|secret|token|api[-_]?key|credential|authorization|private[-_]?key|client[-_]?secret)", Pattern.CASE_INSENSITIVE);

    QueryMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, !id.equals("query.definition.replace"), "edit-cell");
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported query mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        if (id().equals("query.definition.replace")) {
            validateDefinitionChange(root, params);
            return List.of();
        }
        QueryCells payload = cellsPayload(root, params);
        List<RangeRef> ranges = new ArrayList<>();
        ranges.add(payload.clearRange());
        if (payload.pivotSheetId() != null) ranges.add(SnapshotMutationSupport.wholeSheetRange(root, payload.pivotSheetId()));
        return ranges.stream().distinct().toList();
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        if (id().equals("query.definition.replace")) {
            applyDefinition(root, params);
            return root;
        }
        QueryCells payload = cellsPayload(root, params);
        applyDefinition(root, payload.queryId(), payload.definition());
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, payload.clearRange().sheetId());
        SnapshotMutationSupport.clearCells(sheet, payload.clearRange());
        if (payload.previousCells() != null) restorePreviousCells(root, sheet, payload.clearRange(), payload.previousCells());
        else writeValues(sheet, payload.clearRange(), payload.values());
        if (payload.pivotSheetId() != null) updatePivot(root, payload);
        return root;
    }

    private void validateDefinitionChange(ObjectNode root, ObjectNode params) {
        String queryId = SnapshotMutationSupport.text(params, "queryId");
        JsonNode definition = params.get("definition");
        if (definition == null) throw ServiceException.validation("Query definition is required");
        if (!definition.isNull()) validateDefinition(root, queryId, requireObject(definition, "Query definition"));
    }

    private void applyDefinition(ObjectNode root, ObjectNode params) {
        String queryId = SnapshotMutationSupport.text(params, "queryId");
        JsonNode definition = params.get("definition");
        if (definition == null) throw ServiceException.validation("Query definition is required");
        if (definition.isNull()) {
            removeDefinition(root, queryId);
            return;
        }
        ObjectNode object = requireObject(definition, "Query definition");
        validateDefinition(root, queryId, object);
        upsertDefinition(root, queryId, object);
    }

    private void applyDefinition(ObjectNode root, String queryId, JsonNode definition) {
        if (definition == null || definition.isNull()) {
            removeDefinition(root, queryId);
            return;
        }
        ObjectNode object = requireObject(definition, "Query definition");
        validateDefinition(root, queryId, object);
        upsertDefinition(root, queryId, object);
    }

    private QueryCells cellsPayload(ObjectNode root, ObjectNode params) {
        if (!"cells".equals(SnapshotMutationSupport.text(params, "kind"))) throw ServiceException.validation("Query load must be a cells payload");
        String queryId = SnapshotMutationSupport.text(params, "queryId");
        JsonNode definition = params.get("queryDefinition");
        if (definition == null) throw ServiceException.validation("Query load definition is required");
        if (!definition.isNull()) validateDefinition(root, queryId, requireObject(definition, "Query definition"));
        RangeRef clearRange = SnapshotMutationSupport.range(root, params.get("clearRange"));
        if (SnapshotMutationSupport.cellCount(clearRange) > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Query load range is too large");
        ArrayNode values = SnapshotMutationSupport.requiredArray(params, "values");
        if (values.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Query load values are too large");
        int width = 0;
        for (JsonNode row : values) {
            if (!row.isArray()) throw ServiceException.validation("Query load values must contain rows");
            width = Math.max(width, row.size());
            if (row.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Query load value row is too large");
            for (JsonNode cell : row) if (cell != null && !cell.isNull() && !cell.isObject()) throw ServiceException.validation("Query load cells must be objects");
        }
        if (values.size() > clearRange.endRow() - clearRange.startRow() + 1 || width > clearRange.endColumn() - clearRange.startColumn() + 1) {
            throw ServiceException.validation("Query load values do not fit target range");
        }
        JsonNode previous = params.get("previousCells");
        if (previous != null && !previous.isNull() && !previous.isArray()) throw ServiceException.validation("Query previousCells must be an array");
        String pivotSheetId = null;
        String pivotId = null;
        long pivotRevision = 0;
        String pivotRefreshedAt = null;
        JsonNode pivot = params.get("pivot");
        if (pivot != null && !pivot.isNull()) {
            ObjectNode pivotData = requireObject(pivot, "Query pivot refresh");
            pivotSheetId = SnapshotMutationSupport.text(pivotData, "sheetId");
            pivotId = SnapshotMutationSupport.text(pivotData, "pivotId");
            JsonNode revision = pivotData.get("nextRefreshRevision");
            if (revision == null || !revision.isIntegralNumber() || revision.longValue() < 0) throw ServiceException.validation("Query pivot refresh revision is invalid");
            pivotRevision = revision.longValue();
            pivotRefreshedAt = SnapshotMutationSupport.text(pivotData, "nextRefreshedAt");
            SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, pivotSheetId), "pivots"), pivotId, "Pivot");
        }
        return new QueryCells(queryId, definition, clearRange, values, previous == null || previous.isNull() ? null : (ArrayNode) previous, pivotSheetId, pivotId, pivotRevision, pivotRefreshedAt);
    }

    private void restorePreviousCells(ObjectNode root, ObjectNode sheet, RangeRef range, ArrayNode previousCells) {
        if (previousCells.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Query previousCells are too large");
        for (JsonNode raw : previousCells) {
            if (!raw.isObject()) throw ServiceException.validation("Query previous cell must be an object");
            ObjectNode entry = (ObjectNode) raw;
            SnapshotMutationSupport.CellCoordinate coordinate = SnapshotMutationSupport.coordinate(root, range.sheetId(), entry);
            if (!SnapshotMutationSupport.contains(range, coordinate)) throw ServiceException.validation("Query previous cell is outside its range");
            JsonNode value = entry.get("value");
            if (value != null && !value.isNull()) {
                if (!value.isObject()) throw ServiceException.validation("Query previous value must be an object");
                SnapshotMutationSupport.putCell(sheet, coordinate, value);
            }
        }
    }

    private void writeValues(ObjectNode sheet, RangeRef range, ArrayNode values) {
        for (int rowOffset = 0; rowOffset < values.size(); rowOffset++) {
            ArrayNode row = (ArrayNode) values.get(rowOffset);
            for (int columnOffset = 0; columnOffset < row.size(); columnOffset++) {
                JsonNode value = row.get(columnOffset);
                if (value != null && !value.isNull()) {
                    SnapshotMutationSupport.putCell(sheet, new SnapshotMutationSupport.CellCoordinate(range.startRow() + rowOffset, range.startColumn() + columnOffset), value);
                }
            }
        }
    }

    private void updatePivot(ObjectNode root, QueryCells payload) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, payload.pivotSheetId());
        ObjectNode pivot = SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(sheet, "pivots"), payload.pivotId(), "Pivot");
        pivot.put("refreshRevision", payload.pivotRevision());
        pivot.put("lastRefreshedAt", payload.pivotRefreshedAt());
    }

    private void validateDefinition(ObjectNode root, String expectedId, ObjectNode definition) {
        if (!"QueryDefinition".equals(SnapshotMutationSupport.text(definition, "schema"))) throw ServiceException.validation("Query definition schema is invalid");
        if (!expectedId.equals(SnapshotMutationSupport.text(definition, "id"))) throw ServiceException.validation("Query definition id does not match queryId");
        SnapshotMutationSupport.text(definition, "name");
        SnapshotMutationSupport.text(definition, "connectorId");
        ObjectNode config = SnapshotMutationSupport.requiredObject(definition, "connectorConfig");
        validateSafeConfig(config, "connectorConfig");
        JsonNode revision = definition.get("sourceRevision");
        if (revision == null || !revision.isIntegralNumber() || revision.longValue() < 0) throw ServiceException.validation("Query sourceRevision is invalid");
        ArrayNode steps = SnapshotMutationSupport.requiredArray(definition, "steps");
        if (steps.size() > 100) throw ServiceException.validation("Query has too many steps");
        for (JsonNode raw : steps) validateStep(raw);
        JsonNode policy = definition.get("refreshPolicy");
        if (policy != null && !policy.isNull()) validateRefreshPolicy(policy);
    }

    private void validateStep(JsonNode raw) {
        ObjectNode step = requireObject(raw, "Query step");
        SnapshotMutationSupport.text(step, "id");
        if (!STEP_KINDS.contains(SnapshotMutationSupport.text(step, "kind"))) throw ServiceException.validation("Query step kind is invalid");
        SnapshotMutationSupport.text(step, "name");
        SnapshotMutationSupport.requiredObject(step, "config");
        if (!step.path("enabled").isBoolean()) throw ServiceException.validation("Query step enabled must be boolean");
    }

    private void validateRefreshPolicy(JsonNode raw) {
        ObjectNode policy = requireObject(raw, "Query refresh policy");
        String mode = SnapshotMutationSupport.text(policy, "mode");
        if (!Set.of("manual", "on-open", "interval").contains(mode)) throw ServiceException.validation("Query refresh mode is invalid");
        if (mode.equals("interval")) {
            JsonNode interval = policy.get("intervalMs");
            if (interval == null || !interval.isIntegralNumber() || interval.longValue() <= 0) throw ServiceException.validation("Query refresh interval is invalid");
        }
    }

    private void validateSafeConfig(JsonNode value, String path) {
        if (value == null || value.isNull()) return;
        if (value.isArray()) {
            for (int index = 0; index < value.size(); index++) validateSafeConfig(value.get(index), path + "[" + index + "]");
            return;
        }
        if (!value.isObject()) return;
        value.fields().forEachRemaining(entry -> {
            String key = entry.getKey();
            JsonNode child = entry.getValue();
            if (SECRET_KEY.matcher(key).find()) {
                if (!child.isTextual() || !"[redacted]".equals(child.asText())) throw ServiceException.validation("Query connector credential must be redacted at " + path + "." + key);
            } else validateSafeConfig(child, path + "." + key);
        });
    }

    private void upsertDefinition(ObjectNode root, String queryId, ObjectNode definition) {
        ArrayNode definitions = SnapshotMutationSupport.array(root, "queryDefinitions");
        int index = SnapshotMutationSupport.indexById(definitions, queryId);
        if (index >= 0) definitions.set(index, definition.deepCopy());
        else definitions.add(definition.deepCopy());
    }

    private void removeDefinition(ObjectNode root, String queryId) {
        SnapshotMutationSupport.removeById(SnapshotMutationSupport.array(root, "queryDefinitions"), queryId);
    }

    private ObjectNode requireObject(JsonNode value, String label) {
        if (value == null || !value.isObject()) throw ServiceException.validation(label + " must be an object");
        return (ObjectNode) value;
    }

    private record QueryCells(
            String queryId,
            JsonNode definition,
            RangeRef clearRange,
            ArrayNode values,
            ArrayNode previousCells,
            String pivotSheetId,
            String pivotId,
            long pivotRevision,
            String pivotRefreshedAt
    ) {
    }
}
