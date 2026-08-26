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
import java.util.regex.Pattern;

/** Reducers for query definitions and metadata-only block-backed loads. */
final class QueryMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "query.definition.replace", "query.load.range", "query.load.sheet-table",
            "query.load.pivot-source", "query.load.workbook-table"
    );
    private static final Pattern SAFE_ID = Pattern.compile("[A-Za-z0-9._:-]{1,200}");
    private static final Pattern SHA256 = Pattern.compile("[A-Fa-f0-9]{64}");
    private static final Set<String> TARGET_KINDS = Set.of("range", "sheet-table", "pivot-source", "workbook-table");
    private static final Set<String> STEP_KINDS = Set.of("source", "filter", "select-columns", "rename-column", "sort", "group-by", "join", "pivot", "custom");
    private static final Pattern SECRET_KEY = Pattern.compile("(?:pass(word)?|secret|token|api[-_]?key|credential|authorization|private[-_]?key|client[-_]?secret)", Pattern.CASE_INSENSITIVE);

    QueryMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR);
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
        QueryLoad payload = loadPayload(root, mutation, params);
        List<RangeRef> ranges = new ArrayList<>();
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject()) continue;
            for (JsonNode rawRegion : SnapshotMutationSupport.array((ObjectNode) rawSheet, "dataRegions")) {
                if (payload.sourceId().equals(rawRegion.path("sourceId").asText())) {
                    ranges.add(SnapshotMutationSupport.range(root, rawRegion.path("range")));
                }
            }
        }
        if (payload.binding() != null && "sheet-region".equals(payload.binding().path("kind").asText())) {
            ranges.add(SnapshotMutationSupport.range(root, payload.binding().path("region").path("range")));
        }
        return ranges;
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        if (id().equals("query.definition.replace")) {
            applyDefinition(root, params);
            return root;
        }
        QueryLoad payload = loadPayload(root, mutation, params);
        applyDefinition(root, payload.queryId(), payload.definition());
        removeCurrentBinding(root, payload.sourceId());
        if (payload.source() != null) {
            ArrayNode sources = SnapshotMutationSupport.dataModelArray(root, "sources");
            int index = SnapshotMutationSupport.indexById(sources, payload.sourceId());
            if (index >= 0) sources.set(index, payload.source().deepCopy());
            else sources.add(payload.source().deepCopy());
        }
        applyBinding(root, payload);
        if (payload.pivotSource() != null) applyPivotSource(root, payload, payload.pivotSource());
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
        applyDefinition(root, queryId, definition);
    }

    private void applyDefinition(ObjectNode root, String queryId, JsonNode definition) {
        if (definition == null || definition.isNull()) {
            SnapshotMutationSupport.removeById(SnapshotMutationSupport.array(root, "queryDefinitions"), queryId);
            return;
        }
        ObjectNode object = requireObject(definition, "Query definition");
        validateDefinition(root, queryId, object);
        upsertDefinition(root, queryId, object);
    }

    private QueryLoad loadPayload(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        SnapshotMutationSupport.validateKnownKeys(params, Set.of(
                "kind", "queryId", "queryDefinition", "target", "sourceId", "source", "binding", "extent", "pivotSource"
        ), "Query load payload");
        if (!"data-source-load".equals(SnapshotMutationSupport.text(params, "kind"))) throw ServiceException.validation("Query load kind is invalid");
        String queryId = identity(SnapshotMutationSupport.text(params, "queryId"), "Query id");
        String sourceId = identity(SnapshotMutationSupport.text(params, "sourceId"), "Query sourceId");
        if (!sourceId.equals("query:" + queryId)) throw ServiceException.validation("Query sourceId does not match queryId");
        JsonNode definition = params.get("queryDefinition");
        if (definition == null) throw ServiceException.validation("Query load definition is required");
        if (!definition.isNull()) validateDefinition(root, queryId, requireObject(definition, "Query definition"));
        ObjectNode target = requireObject(params.get("target"), "Query load target");
        validateTarget(root, target);
        JsonNode sourceNode = params.get("source");
        ObjectNode source = null;
        if (sourceNode != null && !sourceNode.isNull()) {
            source = DataSourceMutationDescriptor.validateQuerySource(root, mutation.sheetId(), requireObject(sourceNode, "Query data source"));
            if (!sourceId.equals(SnapshotMutationSupport.text(source, "id"))) throw ServiceException.validation("Query data source id does not match sourceId");
            if (source.path("fields").size() < 1) throw ServiceException.validation("Query data source must contain at least one field");
        }
        JsonNode bindingNode = params.get("binding");
        ObjectNode binding = null;
        if (bindingNode != null && !bindingNode.isNull()) binding = validateBinding(root, mutation, sourceId, requireObject(bindingNode, "Query load binding"), source);
        JsonNode extent = params.get("extent");
        if (extent != null && !extent.isNull()) validateExtent(root, requireObject(extent, "Query load extent"));
        if (extent != null && !extent.isNull() && binding != null && "sheet-region".equals(binding.path("kind").asText())) {
            RangeRef range = SnapshotMutationSupport.range(root, binding.path("region").path("range"));
            ObjectNode extentObject = (ObjectNode) extent;
            if (!range.sheetId().equals(SnapshotMutationSupport.text(extentObject, "sheetId"))
                    || nonNegative(extentObject, "rowCount") < range.endRow() + 1
                    || nonNegative(extentObject, "columnCount") < range.endColumn() + 1) {
                throw ServiceException.validation("Query load extent does not contain the bound region");
            }
        }
        JsonNode pivotSourceNode = params.get("pivotSource");
        ObjectNode pivotSource = null;
        if (pivotSourceNode != null && !pivotSourceNode.isNull()) pivotSource = validatePivotSource(root, sourceId, requireObject(pivotSourceNode, "Query Pivot source"));
        return new QueryLoad(queryId, definition, target, sourceId, source, binding, extent == null || extent.isNull() ? null : (ObjectNode) extent, pivotSource);
    }

    private ObjectNode validateBinding(ObjectNode root, OperationMutation mutation, String sourceId, ObjectNode binding, ObjectNode source) {
        String kind = SnapshotMutationSupport.text(binding, "kind");
        if ("sheet-region".equals(kind)) {
            SnapshotMutationSupport.validateKnownKeys(binding, Set.of("kind", "region", "header"), "Query sheet binding");
            ObjectNode region = requireObject(binding.get("region"), "Query sheet region");
            SnapshotMutationSupport.validateKnownKeys(region, Set.of("id", "sourceId", "range", "headerRow", "revision"), "Query sheet region");
            if (!sourceId.equals(SnapshotMutationSupport.text(region, "sourceId"))) throw ServiceException.validation("Query region sourceId does not match sourceId");
            RangeRef range = SnapshotMutationSupport.range(root, region.get("range"));
            if (source != null && source.path("sourceSheetId").isTextual() && !source.path("sourceSheetId").asText().equals(range.sheetId())) throw ServiceException.validation("Query sheet region targets another source sheet");
            if (source != null && source.path("sourceRange").isObject() && !sameRange(range, SnapshotMutationSupport.range(root, source.get("sourceRange")))) throw ServiceException.validation("Query sheet region does not match source range");
            long headerRow = nonNegative(region, "headerRow");
            if (headerRow != range.startRow() || headerRow > range.endRow()) throw ServiceException.validation("Query sheet region headerRow is invalid");
            long regionRevision = nonNegative(region, "revision");
            if (source != null && regionRevision != nonNegative(source, "revision")) throw ServiceException.validation("Query sheet region revision does not match source revision");
            ArrayNode header = SnapshotMutationSupport.requiredArray(binding, "header");
            int fieldCount = source == null ? 0 : source.path("fields").size();
            if (header.size() != fieldCount) throw ServiceException.validation("Query sheet header width does not match source fields");
            for (JsonNode value : header) if (!(value.isNull() || value.isTextual() || value.isNumber() || value.isBoolean())) throw ServiceException.validation("Query sheet header contains a non-scalar value");
            return binding;
        }
        if ("workbook-table".equals(kind)) {
            SnapshotMutationSupport.validateKnownKeys(binding, Set.of("kind", "tableId", "table"), "Query workbook table binding");
            String tableId = identity(SnapshotMutationSupport.text(binding, "tableId"), "Query tableId");
            ObjectNode table = requireObject(binding.get("table"), "Query workbook table");
            validateTable(root, table, source == null ? null : sourceId);
            if (!tableId.equals(SnapshotMutationSupport.text(table, "id"))) throw ServiceException.validation("Query workbook table id does not match tableId");
            if (SnapshotMutationSupport.findById(SnapshotMutationSupport.dataModelArray(root, "tables"), tableId) == null) throw ServiceException.notFound("Workbook table not found: " + tableId);
            return binding;
        }
        throw ServiceException.validation("Query load binding kind is invalid");
    }

    private void validateExtent(ObjectNode root, ObjectNode extent) {
        SnapshotMutationSupport.validateKnownKeys(extent, Set.of("sheetId", "rowCount", "columnCount"), "Query load extent");
        String sheetId = SnapshotMutationSupport.text(extent, "sheetId");
        SnapshotMutationSupport.sheet(root, sheetId);
        if (nonNegative(extent, "rowCount") < 1 || nonNegative(extent, "columnCount") < 1) throw ServiceException.validation("Query load extent must be positive");
    }

    private void validateTarget(ObjectNode root, ObjectNode target) {
        SnapshotMutationSupport.validateKnownKeys(target, Set.of("kind", "sheetId", "range", "tableId", "pivotId"), "Query load target");
        String kind = SnapshotMutationSupport.text(target, "kind");
        if (!TARGET_KINDS.contains(kind)) throw ServiceException.validation("Query load target kind is invalid");
        switch (kind) {
            case "range" -> {
                SnapshotMutationSupport.text(target, "sheetId");
                if (target.has("range")) validateTargetRange(root, target.get("sheetId").asText(), target.get("range"));
            }
            case "sheet-table" -> {
                String sheetId = SnapshotMutationSupport.text(target, "sheetId");
                SnapshotMutationSupport.sheet(root, sheetId);
                String tableId = identity(SnapshotMutationSupport.text(target, "tableId"), "Query sheet tableId");
                if (SnapshotMutationSupport.findById(SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, sheetId), "sheetTables"), tableId) == null) throw ServiceException.notFound("Sheet table not found: " + tableId);
            }
            case "workbook-table" -> {
                String tableId = identity(SnapshotMutationSupport.text(target, "tableId"), "Query workbook tableId");
                if (SnapshotMutationSupport.findById(SnapshotMutationSupport.dataModelArray(root, "tables"), tableId) == null) throw ServiceException.notFound("Workbook table not found: " + tableId);
            }
            case "pivot-source" -> {
                String pivotId = identity(SnapshotMutationSupport.text(target, "pivotId"), "Query pivotId");
                boolean found = false;
                for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
                    if (rawSheet.isObject() && SnapshotMutationSupport.findById(SnapshotMutationSupport.array((ObjectNode) rawSheet, "pivots"), pivotId) != null) {
                        found = true;
                        break;
                    }
                }
                if (!found) throw ServiceException.notFound("Pivot not found: " + pivotId);
                if (target.has("sheetId")) SnapshotMutationSupport.sheet(root, target.get("sheetId").asText());
                if (target.has("range")) validateTargetRange(root, target.path("sheetId").asText(), target.get("range"));
            }
            default -> throw ServiceException.validation("Query load target kind is invalid");
        }
    }

    private void validateTargetRange(ObjectNode root, String sheetId, JsonNode rawRange) {
        ObjectNode range = requireObject(rawRange, "Query target range");
        SnapshotMutationSupport.validateKnownKeys(range, Set.of("startRow", "endRow", "startColumn", "endColumn"), "Query target range");
        SnapshotMutationSupport.sheet(root, sheetId);
        long startRow = nonNegative(range, "startRow");
        long startColumn = nonNegative(range, "startColumn");
        if (startRow > Integer.MAX_VALUE || startColumn > Integer.MAX_VALUE) throw ServiceException.validation("Query target range start is too large");
        if (range.has("endRow") && nonNegative(range, "endRow") < startRow) throw ServiceException.validation("Query target range endRow is invalid");
        if (range.has("endColumn") && nonNegative(range, "endColumn") < startColumn) throw ServiceException.validation("Query target range endColumn is invalid");
    }

    private void removeCurrentBinding(ObjectNode root, String sourceId) {
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject()) continue;
            ObjectNode sheet = (ObjectNode) rawSheet;
            ArrayNode regions = SnapshotMutationSupport.array(sheet, "dataRegions");
            for (int index = regions.size() - 1; index >= 0; index--) {
                JsonNode region = regions.get(index);
                if (sourceId.equals(region.path("sourceId").asText())) {
                    clearQueryRegionCells(sheet, SnapshotMutationSupport.range(root, region.path("range")));
                    regions.remove(index);
                }
            }
        }
        ArrayNode tables = SnapshotMutationSupport.dataModelArray(root, "tables");
        for (JsonNode raw : tables) if (raw.isObject() && sourceId.equals(raw.path("sourceId").asText())) ((ObjectNode) raw).remove("sourceId");
        ArrayNode sources = SnapshotMutationSupport.dataModelArray(root, "sources");
        SnapshotMutationSupport.removeById(sources, sourceId);
    }

    private void applyBinding(ObjectNode root, QueryLoad payload) {
        if (payload.extent() != null) {
            ObjectNode extent = payload.extent();
            ObjectNode extentSheet = SnapshotMutationSupport.sheet(root, extent.path("sheetId").asText());
            extentSheet.put("rowCount", extent.path("rowCount").asInt());
            extentSheet.put("columnCount", extent.path("columnCount").asInt());
        }
        if (payload.binding() == null) return;
        ObjectNode binding = payload.binding();
        if ("sheet-region".equals(binding.path("kind").asText())) {
            ObjectNode region = (ObjectNode) binding.get("region");
            RangeRef range = SnapshotMutationSupport.range(root, region.get("range"));
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, range.sheetId());
            clearQueryRegionCells(sheet, range);
            ArrayNode header = (ArrayNode) binding.get("header");
            for (int column = 0; column < header.size(); column++) {
                JsonNode value = header.get(column);
                if (value != null && !value.isNull()) {
                    ObjectNode cell = JsonNodeFactory.instance.objectNode();
                    cell.set("value", value.deepCopy());
                    SnapshotMutationSupport.putCell(sheet, new SnapshotMutationSupport.CellCoordinate(range.startRow(), range.startColumn() + column), cell);
                }
            }
            SnapshotMutationSupport.array(sheet, "dataRegions").add(region.deepCopy());
            return;
        }
        String tableId = SnapshotMutationSupport.text(binding, "tableId");
        ArrayNode tables = SnapshotMutationSupport.dataModelArray(root, "tables");
        int index = SnapshotMutationSupport.indexById(tables, tableId);
        if (index < 0) throw ServiceException.notFound("Workbook table not found: " + tableId);
        tables.set(index, ((ObjectNode) binding.get("table")).deepCopy());
    }

    private void clearQueryRegionCells(ObjectNode sheet, RangeRef range) {
        ObjectNode cells = SnapshotMutationSupport.cells(sheet);
        for (int row = range.startRow(); row <= range.endRow(); row++) {
            ObjectNode current = SnapshotMutationSupport.cellRow(cells, row, false);
            if (current == null) continue;
            List<String> remove = new ArrayList<>();
            current.fieldNames().forEachRemaining(key -> {
                int column;
                try {
                    column = Integer.parseInt(key);
                } catch (NumberFormatException error) {
                    throw ServiceException.validation("Cell column key is invalid");
                }
                JsonNode cell = current.get(key);
                if (column >= range.startColumn() && column <= range.endColumn()
                        && (cell == null || !cell.isObject() || !cell.has("__cellPatch"))) remove.add(key);
            });
            remove.forEach(current::remove);
            if (current.isEmpty()) cells.remove(Integer.toString(row));
        }
    }

    private boolean sameRange(RangeRef left, RangeRef right) {
        return left.sheetId().equals(right.sheetId())
                && left.startRow() == right.startRow()
                && left.endRow() == right.endRow()
                && left.startColumn() == right.startColumn()
                && left.endColumn() == right.endColumn();
    }

    private ObjectNode validatePivotSource(ObjectNode root, String sourceId, ObjectNode source) {
        String kind = SnapshotMutationSupport.text(source, "kind");
        if ("data-source".equals(kind)) {
            SnapshotMutationSupport.validateKnownKeys(source, Set.of("kind", "dataSourceId"), "Query Pivot data source");
            if (!sourceId.equals(SnapshotMutationSupport.text(source, "dataSourceId"))) throw ServiceException.validation("Query Pivot sourceId does not match query source");
            return source;
        }
        if ("worksheet-range".equals(kind)) {
            SnapshotMutationSupport.validateKnownKeys(source, Set.of("kind", "range"), "Query Pivot worksheet source");
            SnapshotMutationSupport.range(root, source.get("range"));
            return source;
        }
        if ("worksheet-ranges".equals(kind)) {
            SnapshotMutationSupport.validateKnownKeys(source, Set.of("kind", "ranges", "relationships"), "Query Pivot worksheet sources");
            if (!source.path("ranges").isArray() || !source.path("relationships").isArray()) throw ServiceException.validation("Query Pivot worksheet sources are invalid");
            return source;
        }
        if ("table".equals(kind)) {
            SnapshotMutationSupport.validateKnownKeys(source, Set.of("kind", "tableId"), "Query Pivot table source");
            SnapshotMutationSupport.text(source, "tableId");
            return source;
        }
        if ("named-range".equals(kind)) {
            SnapshotMutationSupport.validateKnownKeys(source, Set.of("kind", "name", "sheetId"), "Query Pivot named source");
            SnapshotMutationSupport.text(source, "name");
            return source;
        }
        throw ServiceException.validation("Query Pivot source kind is invalid");
    }

    private void applyPivotSource(ObjectNode root, QueryLoad payload, ObjectNode source) {
        ObjectNode target = requireObject(payload.target(), "Query load target");
        String pivotId = SnapshotMutationSupport.text(target, "pivotId");
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject()) continue;
            ObjectNode pivot = SnapshotMutationSupport.findById(SnapshotMutationSupport.array((ObjectNode) rawSheet, "pivots"), pivotId);
            if (pivot != null) {
                pivot.set("source", source.deepCopy());
                return;
            }
        }
        throw ServiceException.notFound("Pivot not found: " + pivotId);
    }

    private void validateTable(ObjectNode root, ObjectNode table, String sourceId) {
        SnapshotMutationSupport.validateKnownKeys(table, Set.of("id", "name", "sourceId", "sourceSheetId", "sourceRange", "rowCount", "fields", "blockSize", "blocks", "revision"), "Query workbook table");
        identity(SnapshotMutationSupport.text(table, "id"), "Workbook table id");
        if (SnapshotMutationSupport.text(table, "name").length() > 255) throw ServiceException.validation("Workbook table name is too long");
        String tableSourceId = SnapshotMutationSupport.optionalText(table, "sourceId");
        if (sourceId != null && !sourceId.equals(tableSourceId)) throw ServiceException.validation("Query workbook table sourceId is invalid");
        if (nonNegative(table, "rowCount") < 0 || nonNegative(table, "blockSize") < 1 || nonNegative(table, "revision") < 0) throw ServiceException.validation("Query workbook table numeric metadata is invalid");
        if (!table.path("fields").isArray() || !table.path("blocks").isArray()) throw ServiceException.validation("Workbook table fields and blocks are required");
        JsonNode sourceRange = table.get("sourceRange");
        String sourceSheetId = SnapshotMutationSupport.optionalText(table, "sourceSheetId");
        if (sourceRange != null && !sourceRange.isNull()) {
            RangeRef range = SnapshotMutationSupport.range(root, sourceRange);
            if (sourceSheetId != null && !sourceSheetId.equals(range.sheetId())) throw ServiceException.validation("Workbook table source sheet does not match source range");
        } else if (sourceSheetId != null) SnapshotMutationSupport.sheet(root, sourceSheetId);
    }

    private void validateDefinition(ObjectNode root, String expectedId, ObjectNode definition) {
        if (!"QueryDefinition".equals(SnapshotMutationSupport.text(definition, "schema"))) throw ServiceException.validation("Query definition schema is invalid");
        if (!expectedId.equals(SnapshotMutationSupport.text(definition, "id"))) throw ServiceException.validation("Query definition id does not match queryId");
        SnapshotMutationSupport.text(definition, "name");
        SnapshotMutationSupport.text(definition, "connectorId");
        validateSafeConfig(SnapshotMutationSupport.requiredObject(definition, "connectorConfig"), "connectorConfig");
        if (nonNegative(definition, "sourceRevision") < 0) throw ServiceException.validation("Query sourceRevision is invalid");
        ArrayNode steps = SnapshotMutationSupport.requiredArray(definition, "steps");
        if (steps.size() > 100) throw ServiceException.validation("Query has too many steps");
        for (JsonNode raw : steps) {
            ObjectNode step = requireObject(raw, "Query step");
            SnapshotMutationSupport.text(step, "id");
            if (!STEP_KINDS.contains(SnapshotMutationSupport.text(step, "kind"))) throw ServiceException.validation("Query step kind is invalid");
            SnapshotMutationSupport.text(step, "name");
            SnapshotMutationSupport.requiredObject(step, "config");
            if (!step.path("enabled").isBoolean()) throw ServiceException.validation("Query step enabled must be boolean");
        }
        JsonNode policy = definition.get("refreshPolicy");
        if (policy != null && !policy.isNull()) {
            ObjectNode value = requireObject(policy, "Query refresh policy");
            String mode = SnapshotMutationSupport.text(value, "mode");
            if (!Set.of("manual", "on-open", "interval").contains(mode)) throw ServiceException.validation("Query refresh mode is invalid");
            if (mode.equals("interval") && nonNegative(value, "intervalMs") <= 0) throw ServiceException.validation("Query refresh interval is invalid");
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
            if (SECRET_KEY.matcher(entry.getKey()).find()) {
                if (!entry.getValue().isTextual() || !"[redacted]".equals(entry.getValue().asText())) throw ServiceException.validation("Query connector credential must be redacted at " + path + "." + entry.getKey());
            } else validateSafeConfig(entry.getValue(), path + "." + entry.getKey());
        });
    }

    private void upsertDefinition(ObjectNode root, String queryId, ObjectNode definition) {
        ArrayNode definitions = SnapshotMutationSupport.array(root, "queryDefinitions");
        int index = SnapshotMutationSupport.indexById(definitions, queryId);
        if (index >= 0) definitions.set(index, definition.deepCopy()); else definitions.add(definition.deepCopy());
    }

    private ObjectNode requireObject(JsonNode value, String label) {
        if (value == null || !value.isObject()) throw ServiceException.validation(label + " must be an object");
        return (ObjectNode) value;
    }

    private static long nonNegative(ObjectNode object, String property) {
        JsonNode value = object.get(property);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToLong() || value.longValue() < 0) throw ServiceException.validation(property + " must be a non-negative integer");
        return value.longValue();
    }

    private static String identity(String value, String label) {
        if (!SAFE_ID.matcher(value).matches()) throw ServiceException.validation(label + " is invalid");
        return value;
    }

    private record QueryLoad(String queryId, JsonNode definition, ObjectNode target, String sourceId, ObjectNode source, ObjectNode binding, ObjectNode extent, ObjectNode pivotSource) {
    }
}
