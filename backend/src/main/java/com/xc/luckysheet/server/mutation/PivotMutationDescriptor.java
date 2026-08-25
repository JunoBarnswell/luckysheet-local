package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;
import com.xc.luckysheet.server.mutation.SnapshotMutationSupport.CellCoordinate;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;

/** Reducers for the persisted PivotDefinition contract; result trees remain derived. */
final class PivotMutationDescriptor extends CanonicalJsonMutationDescriptor {
    private static final int MAX_MEMBER_COUNT = 1_048_576;
    private static final String SCHEMA = "PivotDefinition";
    private static final String FIELD_CATALOG_SCHEMA = "PivotFieldCatalog";
    private static final Set<String> PIVOT_KEYS = Set.of(
            "schema", "id", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "presentation", "nativeMetadata"
    );
    private static final Set<String> UPDATE_KEYS = Set.of(
            "sheetId", "pivotId", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "presentation", "nativeMetadata"
    );
    private static final Set<String> SOURCE_KEYS = Set.of("kind", "range", "ranges", "relationships", "tableId", "name", "sheetId", "dataSourceId");
    private static final Set<String> TARGET_KEYS = Set.of("sheetId", "anchor");
    private static final Set<String> FIELD_KEYS = Set.of("fieldId", "name", "dataType", "ordinal", "values");
    private static final Set<String> LAYOUT_KEYS = Set.of(
            "rows", "columns", "filters", "allowMultipleFiltersPerField", "collation", "values", "calculatedFields", "calculatedItems",
            "subtotalLocation", "showRowGrandTotals", "showColumnGrandTotals", "reportLayout", "expansion"
    );
    private static final Set<String> COLLATION_KEYS = Set.of("locale", "sensitivity", "numeric", "caseFirst");
    private static final Set<String> PLACEMENT_KEYS = Set.of("fieldId", "sort", "group", "subtotal");
    private static final Set<String> VALUE_KEYS = Set.of(
            "valueId", "fieldId", "summarizeBy", "displayName", "numberFormat", "baseFieldId", "baseItem", "showAs"
    );
    private static final Set<String> REFRESH_KEYS = Set.of("mode", "preserveFormatting", "refreshOnLoad");
    private static final Set<String> NATIVE_KEYS = Set.of(
            "cacheId", "cacheDefinitionPart", "cacheRecordsPart", "pivotTablePart", "fieldBindings", "preservedFeatures"
    );
    private static final Set<String> PRESENTATION_KEYS = Set.of("styleName", "styleOptions");
    private static final Set<String> STYLE_OPTIONS_KEYS = Set.of(
            "showRowHeaders", "showColumnHeaders", "showRowStripes", "showColumnStripes", "showLastColumn"
    );
    private static final Set<String> AGGREGATORS = Set.of(
            "sum", "count", "count-numbers", "average", "min", "max", "product", "stdev", "stdevp", "var", "varp", "distinct-count"
    );
    private static final Set<String> FIELD_TYPES = Set.of("text", "number", "date", "boolean", "mixed");
    private static final Set<String> FILTER_KINDS = Set.of("manual", "condition", "top-items");
    private static final Set<String> FILTER_SCOPES = Set.of("report", "field");
    private static final Set<String> MANUAL_MODES = Set.of("all", "include", "exclude");
    private static final Set<String> LABEL_FILTER_OPERATORS = Set.of(
            "equals", "not-equals", "begins-with", "not-begins-with", "ends-with", "not-ends-with",
            "contains", "not-contains", "between", "not-between", "greater-than", "greater-or-equal",
            "less-than", "less-or-equal"
    );
    private static final Set<String> DATE_FILTER_OPERATORS = Set.of("equals", "not-equals", "before", "after", "between", "not-between");
    private static final Set<String> VALUE_FILTER_OPERATORS = Set.of(
            "equals", "not-equals", "greater-than", "greater-or-equal", "less-than", "less-or-equal", "between", "not-between"
    );
    private static final Set<String> DYNAMIC_DATE_FILTERS = Set.of(
            "today", "yesterday", "tomorrow", "this-week", "last-week", "next-week", "this-month", "last-month", "next-month",
            "this-quarter", "last-quarter", "next-quarter", "this-year", "last-year", "next-year", "year-to-date"
    );
    private static final Set<String> SHOW_AS_KINDS = Set.of(
            "normal", "grand-percentage", "row-percentage", "column-percentage", "parent-percentage",
            "difference", "percentage-difference", "running-total", "rank", "index"
    );
    static final Set<String> IDS = Set.of("pivot.add", "pivot.remove", "pivot.update", "pivot.refresh");

    PivotMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, true, "structure");
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported pivot mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = canonicalSnapshot(snapshot);
        ObjectNode params = mutation.id().equals("pivot.remove") ? null : SnapshotMutationSupport.params(mutation);
        return switch (id()) {
            case "pivot.add" -> {
                validatePivot(root, mutation.sheetId(), params);
                yield pivotRanges(root, params);
            }
            case "pivot.remove" -> {
                String id = pivotId(mutation.params());
                ObjectNode current = currentPivot(root, mutation.sheetId(), id);
                validatePivot(root, mutation.sheetId(), current);
                validateNoPivotDependencies(root, id);
                yield pivotRanges(root, current);
            }
            case "pivot.update" -> {
                validateIdParams(params, UPDATE_KEYS, "pivot.update");
                ObjectNode current = currentPivot(root, mutation.sheetId(), pivotId(params));
                ObjectNode next = updatedPivot(current, params);
                validatePivot(root, mutation.sheetId(), next);
                yield pivotRanges(root, next);
            }
            case "pivot.refresh" -> {
                validateIdParams(params, Set.of("sheetId", "pivotId"), "pivot.refresh");
                ObjectNode current = currentPivot(root, mutation.sheetId(), pivotId(params));
                validatePivot(root, mutation.sheetId(), current);
                yield pivotRanges(root, current);
            }
            default -> throw ServiceException.validation("Unsupported pivot mutation: " + id());
        };
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = canonicalSnapshot(snapshot);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        ArrayNode pivots = SnapshotMutationSupport.array(sheet, "pivots");
        ObjectNode params = mutation.id().equals("pivot.remove") ? null : SnapshotMutationSupport.params(mutation);
        switch (id()) {
            case "pivot.add" -> add(root, mutation.sheetId(), pivots, params);
            case "pivot.remove" -> remove(root, pivots, pivotId(mutation.params()));
            case "pivot.update" -> update(root, mutation.sheetId(), pivots, params);
            case "pivot.refresh" -> refresh(root, mutation.sheetId(), params);
            default -> throw ServiceException.validation("Unsupported pivot mutation: " + id());
        }
        return root;
    }

    private void add(ObjectNode root, String sheetId, ArrayNode pivots, ObjectNode pivot) {
        canonicalizeFilterScopes(pivot);
        validatePivot(root, sheetId, pivot);
        String id = SnapshotMutationSupport.text(pivot, "id");
        for (JsonNode sheet : SnapshotMutationSupport.sheets(root)) {
            if (!sheet.isObject()) continue;
            if (SnapshotMutationSupport.findById(SnapshotMutationSupport.array((ObjectNode) sheet, "pivots"), id) != null) {
                throw ServiceException.conflict("Pivot already exists: " + id);
            }
        }
        pivots.add(pivot.deepCopy());
    }

    private void remove(ObjectNode root, ArrayNode pivots, String id) {
        validateNoPivotDependencies(root, id);
        if (!SnapshotMutationSupport.removeById(pivots, id)) throw ServiceException.notFound("Pivot not found: " + id);
    }

    private void validateNoPivotDependencies(ObjectNode root, String pivotId) {
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject()) continue;
            ObjectNode sheet = (ObjectNode) rawSheet;
            ArrayNode drawings = SnapshotMutationSupport.array(sheet, "drawings");
            ObjectNode payloads = SnapshotMutationSupport.object(sheet, "drawingPayloads");
            for (JsonNode rawDrawing : drawings) {
                if (!rawDrawing.isObject()) throw ServiceException.validation("Drawing dependency record is invalid");
                String drawingId = rawDrawing.path("id").asText();
                String payloadId = rawDrawing.path("payloadId").asText();
                JsonNode payload = payloads.get(payloadId);
                if (payload == null || !payload.isObject()) throw ServiceException.validation("Drawing payload is missing: " + payloadId);
                String kind = payload.path("kind").asText();
                boolean primary = ("chart".equals(kind) || "slicer".equals(kind) || "timeline".equals(kind))
                        && pivotId.equals(payload.path("pivotId").asText());
                boolean connected = ("slicer".equals(kind) || "timeline".equals(kind))
                        && arrayContainsConnectionPivot(payload.get("connections"), pivotId);
                if (primary || connected) throw ServiceException.conflict("Pivot has dependent drawing: " + drawingId);
            }
        }
    }

    private boolean arrayContainsConnectionPivot(JsonNode value, String expected) {
        if (value == null || !value.isArray()) return false;
        for (JsonNode entry : value) if (entry.isObject() && expected.equals(entry.path("pivotId").asText())) return true;
        return false;
    }

    private void update(ObjectNode root, String sheetId, ArrayNode pivots, ObjectNode params) {
        validateIdParams(params, UPDATE_KEYS, "pivot.update");
        ObjectNode current = SnapshotMutationSupport.requireById(pivots, pivotId(params), "Pivot");
        ObjectNode next = updatedPivot(current, params);
        canonicalizeFilterScopes(next);
        validatePivot(root, sheetId, next);
        pivots.set(SnapshotMutationSupport.indexById(pivots, pivotId(params)), next);
    }

    /** Refresh is a derived computation request; no refresh status is persisted in the snapshot. */
    private void refresh(ObjectNode root, String sheetId, ObjectNode params) {
        validateIdParams(params, Set.of("sheetId", "pivotId"), "pivot.refresh");
        ObjectNode pivot = currentPivot(root, sheetId, pivotId(params));
        validatePivot(root, sheetId, pivot);
    }

    private ObjectNode updatedPivot(ObjectNode current, ObjectNode params) {
        ObjectNode next = current.deepCopy();
        for (String property : List.of("source", "target", "fieldCatalog", "layout", "refreshPolicy", "presentation", "nativeMetadata")) {
            JsonNode value = params.get(property);
            if (value != null) next.set(property, value.deepCopy());
        }
        return next;
    }

    private ObjectNode currentPivot(ObjectNode root, String sheetId, String id) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        return SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(sheet, "pivots"), id, "Pivot");
    }

    private String pivotId(ObjectNode params) {
        return SnapshotMutationSupport.text(params, "pivotId");
    }

    private String pivotId(JsonNode params) {
        if (params == null || !params.isTextual() || params.asText().isBlank()) throw ServiceException.validation("Pivot id is required");
        return params.asText();
    }

    private void validateIdParams(ObjectNode params, Set<String> allowed, String mutation) {
        SnapshotMutationSupport.validateKnownKeys(params, allowed, mutation + " params");
        SnapshotMutationSupport.text(params, "pivotId");
    }

    private void validatePivot(ObjectNode root, String expectedSheetId, ObjectNode pivot) {
        SnapshotMutationSupport.validateKnownKeys(pivot, PIVOT_KEYS, "Pivot");
        if (!SCHEMA.equals(SnapshotMutationSupport.text(pivot, "schema"))) throw ServiceException.validation("Pivot schema must be PivotDefinition");
        SnapshotMutationSupport.text(pivot, "id");
        ObjectNode target = SnapshotMutationSupport.requiredObject(pivot, "target");
        validateTarget(root, expectedSheetId, target);
        PivotSourceResolver.Resolution source = validateSource(root, pivot.get("source"));
        Set<String> fieldIds = validateFieldCatalog(pivot.get("fieldCatalog"));
        validateSourceFieldCatalog(pivot, fieldIds, source.fieldIds());
        Set<String> effectiveFieldIds = validateCalculatedDefinitions(pivot.get("layout"), fieldIds);
        validateLayout(root, pivot.get("layout"), effectiveFieldIds);
        validateRefreshPolicy(pivot.get("refreshPolicy"));
        validatePresentation(pivot.get("presentation"));
        validateNativeMetadata(pivot.get("nativeMetadata"));
    }

    private void validateTarget(ObjectNode root, String expectedSheetId, ObjectNode target) {
        SnapshotMutationSupport.validateKnownKeys(target, TARGET_KEYS, "Pivot target");
        String targetSheetId = SnapshotMutationSupport.text(target, "sheetId");
        if (!expectedSheetId.equals(targetSheetId)) throw ServiceException.validation("Pivot target sheetId does not match mutation sheetId");
        ObjectNode anchor = SnapshotMutationSupport.requiredObject(target, "anchor");
        coordinate(root, targetSheetId, anchor, "Pivot target anchor");
    }

    private PivotSourceResolver.Resolution validateSource(ObjectNode root, JsonNode raw) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot source is required");
        ObjectNode source = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(source, SOURCE_KEYS, "Pivot source");
        String kind = SnapshotMutationSupport.text(source, "kind");
        switch (kind) {
            case "worksheet-range" -> {
                requireOnly(source, Set.of("kind", "range"), "Pivot worksheet range source");
                SnapshotMutationSupport.range(root, source.get("range"));
            }
            case "worksheet-ranges" -> {
                requireOnly(source, Set.of("kind", "ranges", "relationships"), "Pivot worksheet ranges source");
                ArrayNode ranges = SnapshotMutationSupport.requiredArray(source, "ranges");
                if (ranges.isEmpty() || ranges.size() > 10_000) throw ServiceException.validation("Pivot source ranges are invalid");
                Map<String, RangeRef> sourceRanges = new LinkedHashMap<>();
                for (JsonNode rawRange : ranges) {
                    if (!rawRange.isObject()) throw ServiceException.validation("Pivot source range must be an object");
                    ObjectNode sourceRange = (ObjectNode) rawRange;
                    SnapshotMutationSupport.validateKnownKeys(sourceRange, Set.of("sourceId", "range"), "Pivot source range");
                    String sourceId = SnapshotMutationSupport.text(sourceRange, "sourceId");
                    if (sourceRanges.containsKey(sourceId)) throw ServiceException.validation("Pivot sourceId is duplicated: " + sourceId);
                    sourceRanges.put(sourceId, SnapshotMutationSupport.range(root, sourceRange.get("range")));
                }
                validateRelationships(root, source.get("relationships"), sourceRanges);
            }
            case "table" -> {
                requireOnly(source, Set.of("kind", "tableId"), "Pivot table source");
                SnapshotMutationSupport.text(source, "tableId");
            }
            case "named-range" -> {
                requireOnly(source, source.has("sheetId") ? Set.of("kind", "name", "sheetId") : Set.of("kind", "name"), "Pivot named range source");
                SnapshotMutationSupport.text(source, "name");
                if (source.has("sheetId")) SnapshotMutationSupport.text(source, "sheetId");
            }
            case "data-source" -> {
                requireOnly(source, Set.of("kind", "dataSourceId"), "Pivot data source");
                SnapshotMutationSupport.text(source, "dataSourceId");
            }
            default -> throw ServiceException.validation("Pivot source kind is invalid");
        }
        return PivotSourceResolver.resolve(root, source);
    }

    private void validateRelationships(ObjectNode root, JsonNode raw, Map<String, RangeRef> sourceRanges) {
        if (raw == null || !raw.isArray()) throw ServiceException.validation("Pivot source relationships must be an array");
        ArrayNode relationships = (ArrayNode) raw;
        if (relationships.size() > 10_000) throw ServiceException.validation("Pivot source relationships are invalid");
        Set<String> sourceIds = sourceRanges.keySet();
        Set<String> relationshipIds = new LinkedHashSet<>();
        Set<String> incomingLeft = new LinkedHashSet<>();
        Map<String, String> parent = new HashMap<>();
        Map<String, Set<String>> graph = new HashMap<>();
        sourceIds.forEach(sourceId -> { parent.put(sourceId, sourceId); graph.put(sourceId, new LinkedHashSet<>()); });
        boolean hasLeftJoin = false;
        for (JsonNode value : relationships) {
            if (!value.isObject()) throw ServiceException.validation("Pivot source relationship must be an object");
            ObjectNode relationship = (ObjectNode) value;
            SnapshotMutationSupport.validateKnownKeys(relationship, Set.of("id", "left", "right", "join"), "Pivot source relationship");
            String relationshipId = SnapshotMutationSupport.text(relationship, "id");
            if (!relationshipIds.add(relationshipId)) throw ServiceException.validation("Pivot relationship id is duplicated: " + relationshipId);
            ObjectNode left = validateEndpoint(root, relationship.get("left"), sourceRanges);
            ObjectNode right = validateEndpoint(root, relationship.get("right"), sourceRanges);
            if (left.path("sourceId").asText().equals(right.path("sourceId").asText())) throw ServiceException.validation("Pivot relationship cannot connect a source to itself");
            String join = SnapshotMutationSupport.text(relationship, "join");
            if (!Set.of("inner", "left").contains(join)) {
                throw ServiceException.validation("Pivot source relationship join is invalid");
            }
            String leftId = left.path("sourceId").asText();
            String rightId = right.path("sourceId").asText();
            String leftType = fieldType(root, sourceRanges.get(leftId), leftId, left.path("fieldId").asText());
            String rightType = fieldType(root, sourceRanges.get(rightId), rightId, right.path("fieldId").asText());
            if ("mixed".equals(leftType) || "mixed".equals(rightType) || !leftType.equals(rightType)) {
                throw ServiceException.validation("Pivot relationship key types are incompatible: " + relationshipId);
            }
            assertUniqueLookupKeys(root, sourceRanges.get(rightId), right.path("fieldId").asText(), rightId);
            if ("inner".equals(join)) assertUniqueLookupKeys(root, sourceRanges.get(leftId), left.path("fieldId").asText(), leftId);
            if ("left".equals(join)) { hasLeftJoin = true; incomingLeft.add(right.path("sourceId").asText()); }
            graph.get(leftId).add(rightId);
            graph.get(rightId).add(leftId);
            String leftRoot = findRoot(parent, leftId);
            String rightRoot = findRoot(parent, rightId);
            if (leftRoot.equals(rightRoot)) throw ServiceException.validation("Pivot relationship graph contains a cycle");
            parent.put(leftRoot, rightRoot);
        }
        if (sourceIds.size() > 1 && relationships.isEmpty()) throw ServiceException.validation("Pivot relationship graph is disconnected");
        if (sourceIds.size() > 1) {
            Set<String> roots = new LinkedHashSet<>(sourceIds);
            if (hasLeftJoin) roots.removeAll(incomingLeft);
            else roots = Set.of(sourceIds.stream().sorted().findFirst().orElseThrow());
            if (roots.size() != 1) throw ServiceException.validation("Pivot relationship graph has an ambiguous root");
            Set<String> visited = new LinkedHashSet<>();
            List<String> pending = new ArrayList<>();
            pending.add(sourceIds.iterator().next());
            while (!pending.isEmpty()) {
                String current = pending.remove(pending.size() - 1);
                if (!visited.add(current)) continue;
                pending.addAll(graph.getOrDefault(current, Set.of()));
            }
            if (visited.size() != sourceIds.size()) throw ServiceException.validation("Pivot relationship graph is disconnected");
        }
    }

    private String findRoot(Map<String, String> parent, String sourceId) {
        String current = parent.get(sourceId);
        if (current == null || current.equals(sourceId)) return sourceId;
        String root = findRoot(parent, current);
        parent.put(sourceId, root);
        return root;
    }

    private ObjectNode validateEndpoint(ObjectNode root, JsonNode raw, Map<String, RangeRef> sourceRanges) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot source relationship endpoint is required");
        ObjectNode endpoint = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(endpoint, Set.of("sourceId", "fieldId"), "Pivot source relationship endpoint");
        String sourceId = SnapshotMutationSupport.text(endpoint, "sourceId");
        RangeRef range = sourceRanges.get(sourceId);
        if (range == null) throw ServiceException.validation("Pivot relationship sourceId is unknown: " + sourceId);
        String fieldId = SnapshotMutationSupport.text(endpoint, "fieldId");
        fieldOrdinal(sourceId, range, fieldId);
        return endpoint;
    }

    private int fieldOrdinal(String sourceId, RangeRef range, String fieldId) {
        String prefix = "source:" + sourceId + ":column:";
        if (!fieldId.startsWith(prefix)) throw ServiceException.validation("Pivot relationship field is not owned by source: " + fieldId);
        String ordinalText = fieldId.substring(prefix.length());
        int ordinal;
        try {
            ordinal = Integer.parseInt(ordinalText);
        } catch (NumberFormatException exception) {
            throw ServiceException.validation("Pivot relationship field ordinal is invalid: " + fieldId);
        }
        if (ordinal < 0 || ordinal > range.endColumn() - range.startColumn()) {
            throw ServiceException.validation("Pivot relationship field is outside source range: " + fieldId);
        }
        return ordinal;
    }

    private String fieldType(ObjectNode root, RangeRef range, String sourceId, String fieldId) {
        int ordinal = fieldOrdinal(sourceId, range, fieldId);
        Set<String> types = new LinkedHashSet<>();
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, range.sheetId());
        for (int row = range.startRow() + 1; row <= range.endRow(); row++) {
            ObjectNode cell = SnapshotMutationSupport.cell(sheet, new CellCoordinate(row, range.startColumn() + ordinal), false);
            if (cell == null) continue;
            JsonNode value = cell.get("formulaValue");
            if (value == null || value.isNull()) value = cell.get("value");
            String type = scalarType(value);
            if (type != null) types.add(type);
        }
        if (types.isEmpty() || types.size() > 1) return "mixed";
        return types.iterator().next();
    }

    private String scalarType(JsonNode value) {
        if (value == null || value.isNull() || (value.isTextual() && value.asText().isEmpty())) return null;
        if (value.isBoolean()) return "boolean";
        if (value.isNumber()) return "number";
        if (value.isTextual()) {
            String text = value.asText();
            if (text.matches("\\d{4}-\\d{2}-\\d{2}(?:[T ].*)?")) return "date";
            return "text";
        }
        return "mixed";
    }

    private void assertUniqueLookupKeys(ObjectNode root, RangeRef range, String fieldId, String sourceId) {
        int ordinal = fieldOrdinal(sourceId, range, fieldId);
        Set<String> keys = new LinkedHashSet<>();
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, range.sheetId());
        for (int row = range.startRow() + 1; row <= range.endRow(); row++) {
            ObjectNode cell = SnapshotMutationSupport.cell(sheet, new CellCoordinate(row, range.startColumn() + ordinal), false);
            JsonNode value = cell == null ? null : cell.get("formulaValue");
            if (value == null || value.isNull()) value = cell == null ? null : cell.get("value");
            String key = typedKey(value);
            if (!keys.add(key)) throw ServiceException.validation("Pivot relationship lookup key is not unique: " + sourceId + ":" + fieldId);
        }
    }

    private String typedKey(JsonNode value) {
        if (value == null || value.isNull() || (value.isTextual() && value.asText().isEmpty())) return "blank:";
        if (value.isBoolean()) return "boolean:" + value.asBoolean();
        if (value.isNumber()) return "number:" + value.asText();
        if (value.isTextual()) return "text:" + value.asText();
        return "mixed:" + value.toString();
    }

    private Set<String> validateFieldCatalog(JsonNode raw) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot fieldCatalog is required");
        ObjectNode catalog = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(catalog, Set.of("schema", "fields"), "Pivot fieldCatalog");
        JsonNode schema = catalog.get("schema");
        if (schema != null && (!schema.isTextual() || !FIELD_CATALOG_SCHEMA.equals(schema.asText()))) {
            throw ServiceException.validation("Pivot fieldCatalog schema is invalid");
        }
        ArrayNode fields = SnapshotMutationSupport.requiredArray(catalog, "fields");
        if (fields.size() > 10_000) throw ServiceException.validation("Pivot fieldCatalog is too large");
        Set<String> ids = new LinkedHashSet<>();
        for (int fieldIndex = 0; fieldIndex < fields.size(); fieldIndex++) {
            JsonNode rawField = fields.get(fieldIndex);
            if (!rawField.isObject()) throw ServiceException.validation("Pivot field must be an object");
            ObjectNode field = (ObjectNode) rawField;
            SnapshotMutationSupport.validateKnownKeys(field, FIELD_KEYS, "Pivot field");
            String id = SnapshotMutationSupport.text(field, "fieldId");
            if (!ids.add(id)) throw ServiceException.validation("Pivot fieldId is duplicated: " + id);
            SnapshotMutationSupport.text(field, "name");
            if (!FIELD_TYPES.contains(SnapshotMutationSupport.text(field, "dataType"))) throw ServiceException.validation("Pivot field dataType is invalid");
            JsonNode ordinal = field.get("ordinal");
            if (ordinal == null || !ordinal.isIntegralNumber() || ordinal.intValue() != fieldIndex) throw ServiceException.validation("Pivot field ordinal is invalid");
            if (field.has("values")) {
                JsonNode values = field.get("values");
                if (!values.isArray() || values.size() > MAX_MEMBER_COUNT) {
                    throw ServiceException.validation("Pivot field values are invalid");
                }
                validateScalars(values, "Pivot field values");
            }
        }
        return ids;
    }

    private void validateSourceFieldCatalog(ObjectNode pivot, Set<String> fieldIds, List<String> sourceFieldIds) {
        // An empty catalogue is retained as an explicitly empty preflight
        // shape for reducers that have not yet materialized live source fields.
        // Once a catalogue is supplied, every source field must be the stable
        // identity produced by the same source resolver; this prevents a
        // source switch from silently retaining fields from the old source.
        if (fieldIds.isEmpty()) return;
        Set<String> expected = new LinkedHashSet<>(sourceFieldIds);
        Set<String> calculated = new LinkedHashSet<>();
        JsonNode layout = pivot.get("layout");
        if (layout != null && layout.isObject()) {
            for (String key : List.of("calculatedFields", "calculatedItems")) {
                JsonNode entries = layout.get(key);
                if (entries == null || !entries.isArray()) continue;
                for (JsonNode entry : entries) if (entry.isObject() && entry.path("fieldId").isTextual()) calculated.add(entry.path("fieldId").asText());
            }
        }
        for (String fieldId : fieldIds) {
            if (!expected.contains(fieldId) && !calculated.contains(fieldId)) {
                throw ServiceException.validation("Pivot fieldId is not owned by its source: " + fieldId);
            }
        }
        if (!expected.isEmpty() && !fieldIds.containsAll(expected)) {
            throw ServiceException.validation("Pivot field catalog is incomplete for its source");
        }
    }

    /**
     * Materializes the one canonical scope value at the persistence boundary.
     * An omitted scope is field-scoped only when its field is placed on rows or
     * columns; all other omitted scopes are report-scoped.
     */
    private void canonicalizeFilterScopes(ObjectNode pivot) {
        JsonNode rawLayout = pivot.get("layout");
        if (rawLayout == null || !rawLayout.isObject()) return;
        ObjectNode layout = (ObjectNode) rawLayout;
        JsonNode rawRows = layout.get("rows");
        JsonNode rawColumns = layout.get("columns");
        JsonNode rawFilters = layout.get("filters");
        if (rawFilters == null || !rawFilters.isArray()) return;
        Set<String> axisFieldIds = new LinkedHashSet<>();
        for (JsonNode placements : new JsonNode[]{rawRows, rawColumns}) {
            if (placements == null || !placements.isArray()) continue;
            for (JsonNode placement : placements) {
                if (placement.isObject() && placement.path("fieldId").isTextual()) axisFieldIds.add(placement.path("fieldId").asText());
            }
        }
        for (JsonNode rawFilter : rawFilters) {
            if (!rawFilter.isObject()) continue;
            ObjectNode filter = (ObjectNode) rawFilter;
            if (!filter.has("scope") && filter.path("fieldId").isTextual()) {
                filter.put("scope", axisFieldIds.contains(filter.path("fieldId").asText()) ? "field" : "report");
            }
        }
    }

    private String filterScope(ObjectNode filter, Set<String> axisFieldIds) {
        String fieldId = SnapshotMutationSupport.text(filter, "fieldId");
        String scope = filter.has("scope")
                ? SnapshotMutationSupport.text(filter, "scope")
                : (axisFieldIds.contains(fieldId) ? "field" : "report");
        if (!FILTER_SCOPES.contains(scope)) throw ServiceException.validation("Pivot filter scope is invalid");
        if ("field".equals(scope) && !axisFieldIds.contains(fieldId)) {
            throw ServiceException.validation("Pivot field filter must target a row or column field");
        }
        return scope;
    }

    private void validateLayout(ObjectNode root, JsonNode raw, Set<String> fieldIds) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot layout is required");
        ObjectNode layout = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(layout, LAYOUT_KEYS, "Pivot layout");
        for (String key : List.of("rows", "columns", "filters", "values")) {
            if (!layout.path(key).isArray()) throw ServiceException.validation("Pivot layout " + key + " must be an array");
        }
        if (!layout.path("allowMultipleFiltersPerField").isBoolean()) {
            throw ServiceException.validation("Pivot layout allowMultipleFiltersPerField must be a boolean");
        }
        validateCollation(layout.get("collation"));
        if (!layout.path("subtotalLocation").isTextual() || !Set.of("top", "bottom", "off").contains(layout.path("subtotalLocation").asText())) {
            throw ServiceException.validation("Pivot layout subtotalLocation is invalid");
        }
        for (String key : List.of("showRowGrandTotals", "showColumnGrandTotals")) {
            if (!layout.path(key).isBoolean()) throw ServiceException.validation("Pivot layout " + key + " must be a boolean");
        }
        JsonNode reportLayout = layout.get("reportLayout");
        if (reportLayout == null || !reportLayout.isTextual()
                || !Set.of("compact", "outline", "tabular").contains(reportLayout.asText())) {
            throw ServiceException.validation("Pivot layout reportLayout is invalid");
        }
        Set<String> axisFieldIds = new LinkedHashSet<>();
        Set<String> valueIds = new LinkedHashSet<>();
        for (JsonNode value : layout.path("values")) validateValue(value, fieldIds, valueIds);
        for (JsonNode placement : layout.path("rows")) {
            validatePlacement(placement, fieldIds, valueIds, "Pivot row field");
            axisFieldIds.add(SnapshotMutationSupport.text((ObjectNode) placement, "fieldId"));
        }
        for (JsonNode placement : layout.path("columns")) {
            validatePlacement(placement, fieldIds, valueIds, "Pivot column field");
            axisFieldIds.add(SnapshotMutationSupport.text((ObjectNode) placement, "fieldId"));
        }
        Set<String> filterIdentities = new LinkedHashSet<>();
        Set<String> filterFields = new LinkedHashSet<>();
        for (JsonNode filter : layout.path("filters")) {
            validateFilter(filter, fieldIds, valueIds);
            ObjectNode object = (ObjectNode) filter;
            String fieldId = SnapshotMutationSupport.text(object, "fieldId");
            String scope = filterScope(object, axisFieldIds);
            String family = SnapshotMutationSupport.text(object, "family");
            if (!filterIdentities.add(fieldId + "|" + scope + "|" + family)) {
                throw ServiceException.validation("Pivot filter family is duplicated: " + fieldId + "|" + scope + "|" + family);
            }
            if (!layout.path("allowMultipleFiltersPerField").asBoolean()
                    && !filterFields.add(fieldId + "|" + scope)) {
                throw ServiceException.validation("Multiple Pivot filters are disabled for " + fieldId + "|" + scope);
            }
        }
        validateExpansion(layout.get("expansion"));
    }

    private void validatePlacement(JsonNode raw, Set<String> fieldIds, Set<String> valueIds, String label) {
        if (!raw.isObject()) throw ServiceException.validation(label + " must be an object");
        ObjectNode placement = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(placement, PLACEMENT_KEYS, label);
        requireField(fieldIds, placement);
        validateSort(placement.get("sort"), valueIds, label + " sort");
        validateGroup(placement.get("group"));
        validateSubtotal(placement.get("subtotal"), label + " subtotal");
    }

    private void validateSubtotal(JsonNode raw, String label) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject()) throw ServiceException.validation(label + " must be an object");
        ObjectNode subtotal = (ObjectNode) raw;
        String mode = SnapshotMutationSupport.text(subtotal, "mode");
        switch (mode) {
            case "automatic", "none" -> SnapshotMutationSupport.validateKnownKeys(subtotal, Set.of("mode"), label);
            case "custom" -> {
                SnapshotMutationSupport.validateKnownKeys(subtotal, Set.of("mode", "functions"), label);
                ArrayNode functions = SnapshotMutationSupport.requiredArray(subtotal, "functions");
                if (functions.isEmpty() || functions.size() > AGGREGATORS.size()) throw ServiceException.validation(label + " functions are invalid");
                Set<String> unique = new LinkedHashSet<>();
                for (JsonNode function : functions) {
                    if (!function.isTextual() || !AGGREGATORS.contains(function.asText()) || !unique.add(function.asText())) {
                        throw ServiceException.validation(label + " functions are invalid");
                    }
                }
            }
            default -> throw ServiceException.validation(label + " mode is invalid");
        }
    }

    private void validateFilter(JsonNode raw, Set<String> fieldIds, Set<String> valueIds) {
        if (!raw.isObject()) throw ServiceException.validation("Pivot filter must be an object");
        ObjectNode filter = (ObjectNode) raw;
        String kind = SnapshotMutationSupport.text(filter, "kind");
        if (!FILTER_KINDS.contains(kind)) throw ServiceException.validation("Pivot filter kind is invalid");
        switch (kind) {
            case "manual" -> {
                SnapshotMutationSupport.validateKnownKeys(filter, Set.of("kind", "family", "fieldId", "scope", "mode", "memberKeys"), "Pivot manual filter");
                if (!"manual".equals(SnapshotMutationSupport.text(filter, "family"))) throw ServiceException.validation("Pivot manual filter family is invalid");
                requireField(fieldIds, filter);
                if (filter.has("scope") && !FILTER_SCOPES.contains(SnapshotMutationSupport.text(filter, "scope"))) throw ServiceException.validation("Pivot manual filter scope is invalid");
                String mode = SnapshotMutationSupport.text(filter, "mode");
                if (!MANUAL_MODES.contains(mode)) throw ServiceException.validation("Pivot manual filter mode is invalid");
                validateMemberValues(filter.get("memberKeys"), "Pivot manual filter memberKeys");
                if ("all".equals(mode) && filter.path("memberKeys").size() != 0) throw ServiceException.validation("Pivot manual filter all mode cannot contain memberKeys");
            }
            case "condition" -> {
                SnapshotMutationSupport.validateKnownKeys(filter, Set.of("kind", "family", "fieldId", "valueId", "scope", "operator", "value", "value2", "dynamic", "wholeDay"), "Pivot condition filter");
                String family = SnapshotMutationSupport.text(filter, "family");
                if (!Set.of("label", "date", "value").contains(family)) throw ServiceException.validation("Pivot condition filter family is invalid");
                requireField(fieldIds, filter);
                if (filter.has("valueId")) {
                    requireValue(valueIds, filter, "valueId");
                    if (!"value".equals(family)) throw ServiceException.validation("Pivot condition valueId requires the value filter family");
                }
                if (filter.has("scope") && !FILTER_SCOPES.contains(SnapshotMutationSupport.text(filter, "scope"))) throw ServiceException.validation("Pivot condition filter scope is invalid");
                Set<String> operators = switch (family) {
                    case "label" -> LABEL_FILTER_OPERATORS;
                    case "date" -> DATE_FILTER_OPERATORS;
                    case "value" -> VALUE_FILTER_OPERATORS;
                    default -> Set.of();
                };
                String operator = SnapshotMutationSupport.text(filter, "operator");
                if (!operators.contains(operator)) {
                    throw ServiceException.validation("Pivot condition filter operator is invalid");
                }
                if (!isScalar(filter.get("value"))) throw ServiceException.validation("Pivot condition filter value is invalid");
                if (filter.has("value2") && !isScalar(filter.get("value2"))) throw ServiceException.validation("Pivot condition filter upper value is invalid");
                if (filter.has("dynamic") && (!"date".equals(family) || !DYNAMIC_DATE_FILTERS.contains(SnapshotMutationSupport.text(filter, "dynamic")))) {
                    throw ServiceException.validation("Pivot dynamic date filter is invalid");
                }
                if (("between".equals(operator) || "not-between".equals(operator))
                        && !filter.has("value2") && !filter.has("dynamic")) {
                    throw ServiceException.validation("Pivot range filter requires two bounds");
                }
                if (filter.has("dynamic") && !Set.of("equals", "between").contains(operator)) {
                    throw ServiceException.validation("Pivot dynamic date operator is invalid");
                }
                if (filter.has("wholeDay") && !filter.get("wholeDay").isBoolean()) throw ServiceException.validation("Pivot condition filter wholeDay is invalid");
            }
            case "top-items" -> {
                SnapshotMutationSupport.validateKnownKeys(filter, Set.of("kind", "family", "fieldId", "scope", "count", "valueId", "direction"), "Pivot top-items filter");
                if (!"top-items".equals(SnapshotMutationSupport.text(filter, "family"))) throw ServiceException.validation("Pivot top-items filter family is invalid");
                requireField(fieldIds, filter);
                if (filter.has("scope") && !FILTER_SCOPES.contains(SnapshotMutationSupport.text(filter, "scope"))) throw ServiceException.validation("Pivot top-items filter scope is invalid");
                JsonNode count = filter.get("count");
                if (count == null || !count.isIntegralNumber() || count.intValue() < 1) throw ServiceException.validation("Pivot top-items count is invalid");
                if (!Set.of("top", "bottom").contains(SnapshotMutationSupport.text(filter, "direction"))) throw ServiceException.validation("Pivot top-items direction is invalid");
                requireValue(valueIds, filter, "valueId");
            }
            default -> throw ServiceException.validation("Pivot filter kind is invalid");
        }
    }

    private void validateCollation(JsonNode raw) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot collation is required");
        ObjectNode collation = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(collation, COLLATION_KEYS, "Pivot collation");
        String locale = SnapshotMutationSupport.text(collation, "locale");
        if (locale.isBlank()) throw ServiceException.validation("Pivot collation locale is invalid");
        if (!Set.of("base", "accent", "case", "variant").contains(SnapshotMutationSupport.text(collation, "sensitivity"))) {
            throw ServiceException.validation("Pivot collation sensitivity is invalid");
        }
        if (!collation.path("numeric").isBoolean()) throw ServiceException.validation("Pivot collation numeric must be a boolean");
        if (!Set.of("upper", "lower", "false").contains(SnapshotMutationSupport.text(collation, "caseFirst"))) {
            throw ServiceException.validation("Pivot collation caseFirst is invalid");
        }
    }

    private void validateValue(JsonNode raw, Set<String> fieldIds, Set<String> valueIds) {
        if (!raw.isObject()) throw ServiceException.validation("Pivot value field must be an object");
        ObjectNode value = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(value, VALUE_KEYS, "Pivot value field");
        JsonNode valueId = value.get("valueId");
        if (valueId == null || !valueId.isTextual() || valueId.asText().isBlank() || !valueIds.add(valueId.asText())) {
            throw ServiceException.validation("Pivot Values placement identity is missing or duplicated");
        }
        requireField(fieldIds, value);
        if (!AGGREGATORS.contains(SnapshotMutationSupport.text(value, "summarizeBy"))) throw ServiceException.validation("Pivot value aggregator is invalid");
        if (value.has("baseFieldId")) requireField(fieldIds, value, "baseFieldId");
        if (value.has("baseItem") && !isScalar(value.get("baseItem")) && !value.get("baseItem").isObject()) throw ServiceException.validation("Pivot baseItem is invalid");
        validateShowAs(value.get("showAs"));
    }

    /**
     * Calculated definitions are layout-owned fields.  A layout-only update
     * therefore cannot require the new definition id to already be present
     * in the persisted source field catalogue.  Register every definition
     * first, then validate item targets against the complete effective set so
     * one update can introduce and reference a calculated field atomically.
     */
    private Set<String> validateCalculatedDefinitions(JsonNode rawLayout, Set<String> catalogFieldIds) {
        Set<String> effectiveFieldIds = new LinkedHashSet<>(catalogFieldIds);
        if (rawLayout == null || !rawLayout.isObject()) return effectiveFieldIds;
        ObjectNode layout = (ObjectNode) rawLayout;
        List<ObjectNode> calculatedItems = new ArrayList<>();
        registerCalculatedDefinitions(layout.get("calculatedFields"), false, effectiveFieldIds);
        collectCalculatedDefinitions(layout.get("calculatedItems"), true, effectiveFieldIds, calculatedItems);
        for (ObjectNode item : calculatedItems) requireField(effectiveFieldIds, item, "targetFieldId");
        return effectiveFieldIds;
    }

    private void registerCalculatedDefinitions(JsonNode raw, boolean item, Set<String> effectiveFieldIds) {
        collectCalculatedDefinitions(raw, item, effectiveFieldIds, null);
    }

    private void collectCalculatedDefinitions(JsonNode raw, boolean item, Set<String> effectiveFieldIds, List<ObjectNode> collected) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isArray() || raw.size() > 10_000) throw ServiceException.validation("Pivot calculated definitions are invalid");
        for (JsonNode value : raw) {
            if (!value.isObject()) throw ServiceException.validation("Pivot calculated definition must be an object");
            ObjectNode definition = (ObjectNode) value;
            SnapshotMutationSupport.validateKnownKeys(definition, item ? Set.of("fieldId", "targetFieldId", "name", "formula") : Set.of("fieldId", "name", "formula"), item ? "Pivot calculated item" : "Pivot calculated field");
            String fieldId = SnapshotMutationSupport.text(definition, "fieldId");
            SnapshotMutationSupport.text(definition, "name");
            SnapshotMutationSupport.text(definition, "formula");
            if (!effectiveFieldIds.add(fieldId)) {
                throw ServiceException.validation("Pivot calculated fieldId is duplicated or collides with the field catalogue: " + fieldId);
            }
            if (item) {
                SnapshotMutationSupport.text(definition, "targetFieldId");
                if (collected != null) collected.add(definition);
            }
        }
    }

    private void validateExpansion(JsonNode raw) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject()) throw ServiceException.validation("Pivot expansion must be an object");
        ObjectNode expansion = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(expansion, Set.of("expandedNodeIds", "collapsedNodeIds", "showButtons"), "Pivot expansion");
        validateStringArray(expansion.get("expandedNodeIds"), "Pivot expandedNodeIds");
        validateStringArray(expansion.get("collapsedNodeIds"), "Pivot collapsedNodeIds");
        if (!expansion.path("showButtons").isBoolean()) throw ServiceException.validation("Pivot expansion showButtons must be a boolean");
    }

    private void validateSort(JsonNode raw, Set<String> valueIds, String label) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject()) throw ServiceException.validation(label + " must be an object");
        ObjectNode sort = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(sort, Set.of("direction", "by", "valueId"), label);
        if (!Set.of("ascending", "descending").contains(SnapshotMutationSupport.text(sort, "direction"))) throw ServiceException.validation(label + " direction is invalid");
        if (sort.has("by") && !Set.of("label", "value").contains(SnapshotMutationSupport.text(sort, "by"))) throw ServiceException.validation(label + " by is invalid");
        if (sort.has("valueId")) requireValue(valueIds, sort, "valueId");
        if ("value".equals(sort.path("by").asText()) && !sort.has("valueId")) {
            throw ServiceException.validation(label + " requires valueId for value sorting");
        }
    }

    private void validateGroup(JsonNode raw) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject()) throw ServiceException.validation("Pivot group must be an object");
        ObjectNode group = (ObjectNode) raw;
        String kind = SnapshotMutationSupport.text(group, "kind");
        switch (kind) {
            case "date" -> {
                SnapshotMutationSupport.validateKnownKeys(group, Set.of("kind", "unit", "startOfWeek"), "Pivot date group");
                if (!Set.of("year", "quarter", "month", "week", "day").contains(SnapshotMutationSupport.text(group, "unit"))) throw ServiceException.validation("Pivot date group unit is invalid");
                if (group.has("startOfWeek") && (!group.get("startOfWeek").isIntegralNumber() || group.get("startOfWeek").intValue() < 0 || group.get("startOfWeek").intValue() > 6)) throw ServiceException.validation("Pivot date group startOfWeek is invalid");
            }
            case "number" -> {
                SnapshotMutationSupport.validateKnownKeys(group, Set.of("kind", "interval", "start", "end"), "Pivot number group");
                if (!finiteNumber(group.get("interval")) || group.path("interval").asDouble() <= 0) throw ServiceException.validation("Pivot number group interval is invalid");
                if (group.has("start") && !finiteNumber(group.get("start"))) throw ServiceException.validation("Pivot number group start is invalid");
                if (group.has("end") && !finiteNumber(group.get("end"))) throw ServiceException.validation("Pivot number group end is invalid");
            }
            case "manual" -> {
                SnapshotMutationSupport.validateKnownKeys(group, Set.of("kind", "groups"), "Pivot manual group");
                ArrayNode groups = SnapshotMutationSupport.requiredArray(group, "groups");
                if (groups.size() > 10_000) throw ServiceException.validation("Pivot manual groups are invalid");
                for (JsonNode rawGroup : groups) {
                    if (!rawGroup.isObject()) throw ServiceException.validation("Pivot manual group entry must be an object");
                    ObjectNode entry = (ObjectNode) rawGroup;
                    SnapshotMutationSupport.validateKnownKeys(entry, Set.of("groupId", "name", "items"), "Pivot manual group entry");
                    SnapshotMutationSupport.text(entry, "groupId");
                    SnapshotMutationSupport.text(entry, "name");
                    validateMemberValues(entry.get("items"), "Pivot manual group items");
                }
            }
            default -> throw ServiceException.validation("Pivot group kind is invalid");
        }
    }

    private void validateShowAs(JsonNode raw) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject()) throw ServiceException.validation("Pivot showAs must be an object");
        ObjectNode showAs = (ObjectNode) raw;
        String kind = SnapshotMutationSupport.text(showAs, "kind");
        if (!SHOW_AS_KINDS.contains(kind)) throw ServiceException.validation("Pivot showAs kind is invalid");
        switch (kind) {
            case "difference", "percentage-difference" -> {
                SnapshotMutationSupport.validateKnownKeys(showAs, Set.of("kind", "base"), "Pivot difference showAs");
                if (!Set.of("grand", "row", "column", "parent").contains(SnapshotMutationSupport.text(showAs, "base"))) throw ServiceException.validation("Pivot showAs base is invalid");
            }
            case "running-total" -> {
                SnapshotMutationSupport.validateKnownKeys(showAs, Set.of("kind", "axis"), "Pivot running-total showAs");
                if (!Set.of("row", "column").contains(SnapshotMutationSupport.text(showAs, "axis"))) throw ServiceException.validation("Pivot showAs axis is invalid");
            }
            case "rank" -> {
                SnapshotMutationSupport.validateKnownKeys(showAs, Set.of("kind", "axis", "direction"), "Pivot rank showAs");
                if (!Set.of("row", "column").contains(SnapshotMutationSupport.text(showAs, "axis")) || !Set.of("ascending", "descending").contains(SnapshotMutationSupport.text(showAs, "direction"))) throw ServiceException.validation("Pivot rank showAs is invalid");
            }
            default -> {
                if (showAs.size() != 1) throw ServiceException.validation("Pivot showAs contains unknown fields");
            }
        }
    }

    private void validateRefreshPolicy(JsonNode raw) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot refreshPolicy is required");
        ObjectNode policy = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(policy, REFRESH_KEYS, "Pivot refreshPolicy");
        if (!Set.of("manual", "on-open", "on-change").contains(SnapshotMutationSupport.text(policy, "mode"))) throw ServiceException.validation("Pivot refreshPolicy mode is invalid");
        if (!policy.path("preserveFormatting").isBoolean() || !policy.path("refreshOnLoad").isBoolean()) throw ServiceException.validation("Pivot refreshPolicy flags are required");
    }

    private void validateNativeMetadata(JsonNode raw) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject()) throw ServiceException.validation("Pivot nativeMetadata must be an object");
        ObjectNode metadata = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(metadata, NATIVE_KEYS, "Pivot nativeMetadata");
        if (metadata.has("cacheId") && (!metadata.get("cacheId").isIntegralNumber() || metadata.get("cacheId").intValue() < 0)) throw ServiceException.validation("Pivot nativeMetadata cacheId is invalid");
        for (String key : List.of("cacheDefinitionPart", "cacheRecordsPart", "pivotTablePart")) if (metadata.has(key)) SnapshotMutationSupport.text(metadata, key);
        if (metadata.has("fieldBindings")) {
            if (!metadata.get("fieldBindings").isObject()) throw ServiceException.validation("Pivot nativeMetadata fieldBindings must be an object");
            ((ObjectNode) metadata.get("fieldBindings")).fields().forEachRemaining(entry -> {
                if (!entry.getValue().isObject()) throw ServiceException.validation("Pivot nativeMetadata field binding must be an object");
                ObjectNode binding = (ObjectNode) entry.getValue();
                SnapshotMutationSupport.validateKnownKeys(binding, Set.of("cacheFieldIndex", "sourceName"), "Pivot nativeMetadata field binding");
                JsonNode index = binding.get("cacheFieldIndex");
                if (index == null || !index.isIntegralNumber() || index.intValue() < 0) throw ServiceException.validation("Pivot nativeMetadata cacheFieldIndex is invalid");
                if (binding.has("sourceName")) SnapshotMutationSupport.text(binding, "sourceName");
            });
        }
        if (metadata.has("preservedFeatures")) {
            ArrayNode features = SnapshotMutationSupport.requiredArray(metadata, "preservedFeatures");
            for (JsonNode feature : features) if (!feature.isTextual() || !Set.of("external-connection", "olap", "consolidation", "macro", "custom-xml", "slicer", "timeline").contains(feature.asText())) throw ServiceException.validation("Pivot nativeMetadata preserved feature is invalid");
        }
    }

    private void validatePresentation(JsonNode raw) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject()) throw ServiceException.validation("Pivot presentation must be an object");
        ObjectNode presentation = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(presentation, PRESENTATION_KEYS, "Pivot presentation");
        if (presentation.has("styleName")) SnapshotMutationSupport.text(presentation, "styleName");
        ObjectNode options = requiredObject(presentation, "styleOptions", "Pivot presentation styleOptions");
        SnapshotMutationSupport.validateKnownKeys(options, STYLE_OPTIONS_KEYS, "Pivot presentation styleOptions");
        for (String key : STYLE_OPTIONS_KEYS) if (!options.path(key).isBoolean()) {
            throw ServiceException.validation("Pivot presentation style option " + key + " must be a boolean");
        }
    }

    private void validateMemberValues(JsonNode raw, String label) {
        if (raw == null || !raw.isArray() || raw.size() > MAX_MEMBER_COUNT) throw ServiceException.validation(label + " must be an array");
        for (JsonNode value : raw) {
            if (!value.isObject()) throw ServiceException.validation(label + " entries must be typed member keys");
            ObjectNode key = (ObjectNode) value;
            SnapshotMutationSupport.validateKnownKeys(key, Set.of("type", "value"), label);
            String type = SnapshotMutationSupport.text(key, "type");
            if (!Set.of("text", "number", "boolean", "blank").contains(type)) throw ServiceException.validation(label + " member type is invalid");
            JsonNode member = key.get("value");
            if ("blank".equals(type)) {
                if (member == null || !member.isNull()) throw ServiceException.validation(label + " blank member must have null value");
            } else if (!isScalar(member) || ("text".equals(type) && !member.isTextual()) || ("number".equals(type) && !member.isNumber()) || ("boolean".equals(type) && !member.isBoolean())) {
                throw ServiceException.validation(label + " member value does not match its type");
            }
        }
    }

    private void validateScalars(JsonNode raw, String label) {
        if (raw == null || !raw.isArray() || raw.size() > MAX_MEMBER_COUNT) throw ServiceException.validation(label + " must be an array");
        for (JsonNode value : raw) if (!isScalar(value)) throw ServiceException.validation(label + " entries must be scalar values");
    }

    private void requireField(Set<String> fieldIds, ObjectNode object) {
        requireField(fieldIds, object, "fieldId");
    }

    private void requireField(Set<String> fieldIds, ObjectNode object, String property) {
        JsonNode value = object.get(property);
        if (value == null || !value.isTextual() || value.asText().isBlank() || !fieldIds.contains(value.asText())) throw ServiceException.validation("Pivot " + property + " is unknown");
    }

    private void requireValue(Set<String> valueIds, ObjectNode object, String property) {
        JsonNode value = object.get(property);
        if (value == null || !value.isTextual() || value.asText().isBlank() || !valueIds.contains(value.asText())) {
            throw ServiceException.validation("Pivot " + property + " is an unknown Values placement");
        }
    }

    private void validateStringArray(JsonNode raw, String label) {
        if (raw == null || !raw.isArray() || raw.size() > 100_000) throw ServiceException.validation(label + " must be an array");
        for (JsonNode value : raw) if (!value.isTextual() || value.asText().isBlank()) throw ServiceException.validation(label + " entries must be non-empty strings");
    }

    private boolean isScalar(JsonNode raw) {
        return raw != null && (raw.isTextual() || raw.isNumber() || raw.isBoolean() || raw.isNull());
    }

    private boolean finiteNumber(JsonNode raw) {
        return raw != null && raw.isNumber() && Double.isFinite(raw.asDouble());
    }

    private void coordinate(ObjectNode root, String sheetId, ObjectNode anchor, String label) {
        SnapshotMutationSupport.validateKnownKeys(anchor, Set.of("row", "column"), label);
        try {
            SnapshotMutationSupport.coordinate(root, sheetId, anchor);
        } catch (ServiceException error) {
            throw ServiceException.validation(label + " is invalid");
        }
    }

    private void requireOnly(ObjectNode object, Set<String> keys, String label) {
        SnapshotMutationSupport.validateKnownKeys(object, keys, label);
    }

    private List<RangeRef> pivotRanges(ObjectNode root, ObjectNode pivot) {
        List<RangeRef> ranges = new ArrayList<>();
        ObjectNode target = SnapshotMutationSupport.requiredObject(pivot, "target");
        String targetSheetId = SnapshotMutationSupport.text(target, "sheetId");
        ranges.add(SnapshotMutationSupport.wholeSheetRange(root, targetSheetId));
        ranges.addAll(sourceRanges(root, pivot));
        return ranges.stream().distinct().toList();
    }

    /** Resolves only canonical source kinds. */
    static List<RangeRef> sourceRanges(ObjectNode root, ObjectNode pivot) {
        ObjectNode source = requiredObject(pivot, "source", "Pivot source");
        return PivotSourceResolver.resolve(root, source).ranges();
    }

    static void forEachWorksheetSourceRange(ObjectNode pivot, Consumer<JsonNode> consumer) {
        JsonNode sourceRaw = pivot.get("source");
        if (sourceRaw == null || !sourceRaw.isObject()) return;
        ObjectNode source = (ObjectNode) sourceRaw;
        switch (source.path("kind").asText()) {
            case "worksheet-range" -> consumer.accept(source.get("range"));
            case "worksheet-ranges" -> {
                for (JsonNode sourceRange : source.path("ranges")) consumer.accept(sourceRange.path("range"));
            }
            default -> {
                // Table, named-range and external data sources are moved by their own participants.
            }
        }
    }

    static ObjectNode requiredTarget(JsonNode pivot) {
        if (pivot == null || !pivot.isObject()) throw ServiceException.validation("Pivot must be an object");
        return requiredObject((ObjectNode) pivot, "target", "Pivot target");
    }

    static ObjectNode requiredAnchor(ObjectNode target) {
        return requiredObject(target, "anchor", "Pivot target anchor");
    }

    private static ObjectNode requiredObject(ObjectNode parent, String property, String label) {
        JsonNode value = parent.get(property);
        if (value == null || !value.isObject()) throw ServiceException.validation(label + " is required");
        return (ObjectNode) value;
    }

    /** Rejects non-canonical persisted records before any reducer reads them. */
    static ObjectNode canonicalSnapshot(JsonNode snapshot) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        assertCanonicalSnapshot(root);
        return root;
    }

    static void assertCanonicalSnapshot(JsonNode snapshot) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject()) throw ServiceException.validation("Workbook contains an invalid sheet");
            ObjectNode sheet = (ObjectNode) rawSheet;
            JsonNode rawPivots = sheet.get("pivots");
            if (rawPivots == null || rawPivots.isNull()) continue;
            if (!rawPivots.isArray()) throw ServiceException.validation("Pivot collection must be an array");
            ArrayNode pivots = (ArrayNode) rawPivots;
            for (JsonNode rawPivot : pivots) {
                if (!rawPivot.isObject()) throw ServiceException.validation("Pivot must be an object");
                ObjectNode pivot = (ObjectNode) rawPivot;
                SnapshotMutationSupport.validateKnownKeys(pivot, PIVOT_KEYS, "Pivot");
                if (!SCHEMA.equals(pivot.path("schema").asText())) throw ServiceException.validation("Pivot schema must be PivotDefinition");
            }
        }
    }
}
