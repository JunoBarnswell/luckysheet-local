package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.function.Consumer;

/** Reducers for the persisted PivotDefinition contract; result trees remain derived. */
final class PivotMutationDescriptor extends CanonicalJsonMutationDescriptor {
    private static final String SCHEMA = "PivotDefinition";
    private static final String FIELD_CATALOG_SCHEMA = "PivotFieldCatalog";
    private static final Set<String> PIVOT_KEYS = Set.of(
            "schema", "id", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "nativeMetadata"
    );
    private static final Set<String> UPDATE_KEYS = Set.of(
            "sheetId", "pivotId", "source", "target", "fieldCatalog", "layout", "refreshPolicy", "nativeMetadata"
    );
    private static final Set<String> SOURCE_KEYS = Set.of("kind", "range", "ranges", "relationships", "tableId", "name", "dataSourceId");
    private static final Set<String> TARGET_KEYS = Set.of("sheetId", "anchor");
    private static final Set<String> FIELD_KEYS = Set.of("fieldId", "name", "dataType", "ordinal", "values");
    private static final Set<String> LAYOUT_KEYS = Set.of(
            "rows", "columns", "filters", "values", "calculatedFields", "calculatedItems",
            "showSubtotals", "showGrandTotals", "compact", "repeatLabels", "expansion"
    );
    private static final Set<String> PLACEMENT_KEYS = Set.of("fieldId", "sort", "group");
    private static final Set<String> VALUE_KEYS = Set.of(
            "fieldId", "summarizeBy", "displayName", "numberFormat", "baseFieldId", "baseItem", "showAs"
    );
    private static final Set<String> REFRESH_KEYS = Set.of("mode", "preserveFormatting", "refreshOnLoad");
    private static final Set<String> NATIVE_KEYS = Set.of(
            "cacheId", "cacheDefinitionPart", "cacheRecordsPart", "pivotTablePart", "fieldBindings", "preservedFeatures"
    );
    private static final Set<String> AGGREGATORS = Set.of(
            "sum", "count", "count-numbers", "average", "min", "max", "product", "stdev", "stdevp", "var", "varp", "distinct-count"
    );
    private static final Set<String> FIELD_TYPES = Set.of("text", "number", "date", "boolean", "mixed");
    private static final Set<String> FILTER_KINDS = Set.of("manual", "condition", "top-items");
    private static final Set<String> MANUAL_MODES = Set.of("all", "include", "exclude");
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
            case "pivot.remove" -> remove(pivots, pivotId(mutation.params()));
            case "pivot.update" -> update(root, mutation.sheetId(), pivots, params);
            case "pivot.refresh" -> refresh(root, mutation.sheetId(), params);
            default -> throw ServiceException.validation("Unsupported pivot mutation: " + id());
        }
        return root;
    }

    private void add(ObjectNode root, String sheetId, ArrayNode pivots, ObjectNode pivot) {
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

    private void remove(ArrayNode pivots, String id) {
        if (!SnapshotMutationSupport.removeById(pivots, id)) throw ServiceException.notFound("Pivot not found: " + id);
    }

    private void update(ObjectNode root, String sheetId, ArrayNode pivots, ObjectNode params) {
        validateIdParams(params, UPDATE_KEYS, "pivot.update");
        ObjectNode current = SnapshotMutationSupport.requireById(pivots, pivotId(params), "Pivot");
        ObjectNode next = updatedPivot(current, params);
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
        for (String property : List.of("source", "target", "fieldCatalog", "layout", "refreshPolicy", "nativeMetadata")) {
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
        validateSource(root, pivot.get("source"));
        Set<String> fieldIds = validateFieldCatalog(pivot.get("fieldCatalog"));
        validateLayout(root, pivot.get("layout"), fieldIds);
        validateRefreshPolicy(pivot.get("refreshPolicy"));
        validateNativeMetadata(pivot.get("nativeMetadata"));
    }

    private void validateTarget(ObjectNode root, String expectedSheetId, ObjectNode target) {
        SnapshotMutationSupport.validateKnownKeys(target, TARGET_KEYS, "Pivot target");
        String targetSheetId = SnapshotMutationSupport.text(target, "sheetId");
        if (!expectedSheetId.equals(targetSheetId)) throw ServiceException.validation("Pivot target sheetId does not match mutation sheetId");
        ObjectNode anchor = SnapshotMutationSupport.requiredObject(target, "anchor");
        coordinate(root, targetSheetId, anchor, "Pivot target anchor");
    }

    private void validateSource(ObjectNode root, JsonNode raw) {
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
                for (JsonNode range : ranges) SnapshotMutationSupport.range(root, range);
                validateRelationships(root, source.get("relationships"));
            }
            case "table" -> {
                requireOnly(source, Set.of("kind", "tableId"), "Pivot table source");
                SnapshotMutationSupport.text(source, "tableId");
            }
            case "named-range" -> {
                requireOnly(source, Set.of("kind", "name"), "Pivot named range source");
                SnapshotMutationSupport.text(source, "name");
            }
            case "data-source" -> {
                requireOnly(source, Set.of("kind", "dataSourceId"), "Pivot data source");
                SnapshotMutationSupport.text(source, "dataSourceId");
            }
            default -> throw ServiceException.validation("Pivot source kind is invalid");
        }
    }

    private void validateRelationships(ObjectNode root, JsonNode raw) {
        if (raw == null || !raw.isArray()) throw ServiceException.validation("Pivot source relationships must be an array");
        ArrayNode relationships = (ArrayNode) raw;
        if (relationships.size() > 10_000) throw ServiceException.validation("Pivot source relationships are invalid");
        for (JsonNode value : relationships) {
            if (!value.isObject()) throw ServiceException.validation("Pivot source relationship must be an object");
            ObjectNode relationship = (ObjectNode) value;
            SnapshotMutationSupport.validateKnownKeys(relationship, Set.of("id", "left", "right", "join"), "Pivot source relationship");
            SnapshotMutationSupport.text(relationship, "id");
            validateEndpoint(root, relationship.get("left"));
            validateEndpoint(root, relationship.get("right"));
            if (!Set.of("inner", "left").contains(SnapshotMutationSupport.text(relationship, "join"))) {
                throw ServiceException.validation("Pivot source relationship join is invalid");
            }
        }
    }

    private void validateEndpoint(ObjectNode root, JsonNode raw) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot source relationship endpoint is required");
        ObjectNode endpoint = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(endpoint, Set.of("sheetId", "fieldId"), "Pivot source relationship endpoint");
        String sheetId = SnapshotMutationSupport.text(endpoint, "sheetId");
        SnapshotMutationSupport.sheet(root, sheetId);
        SnapshotMutationSupport.text(endpoint, "fieldId");
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
            if (field.has("values")) validateScalars(field.get("values"), "Pivot field values");
        }
        return ids;
    }

    private void validateLayout(ObjectNode root, JsonNode raw, Set<String> fieldIds) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Pivot layout is required");
        ObjectNode layout = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(layout, LAYOUT_KEYS, "Pivot layout");
        for (String key : List.of("rows", "columns", "filters", "values")) {
            if (!layout.path(key).isArray()) throw ServiceException.validation("Pivot layout " + key + " must be an array");
        }
        for (String key : List.of("showSubtotals", "showGrandTotals", "compact", "repeatLabels")) {
            if (!layout.path(key).isBoolean()) throw ServiceException.validation("Pivot layout " + key + " must be a boolean");
        }
        for (JsonNode placement : layout.path("rows")) validatePlacement(placement, fieldIds, "Pivot row field");
        for (JsonNode placement : layout.path("columns")) validatePlacement(placement, fieldIds, "Pivot column field");
        for (JsonNode filter : layout.path("filters")) validateFilter(filter, fieldIds);
        for (JsonNode value : layout.path("values")) validateValue(value, fieldIds);
        validateCalculated(layout.get("calculatedFields"), fieldIds, false);
        validateCalculated(layout.get("calculatedItems"), fieldIds, true);
        validateExpansion(layout.get("expansion"));
    }

    private void validatePlacement(JsonNode raw, Set<String> fieldIds, String label) {
        if (!raw.isObject()) throw ServiceException.validation(label + " must be an object");
        ObjectNode placement = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(placement, PLACEMENT_KEYS, label);
        requireField(fieldIds, placement);
        validateSort(placement.get("sort"), fieldIds, label + " sort");
        validateGroup(placement.get("group"));
    }

    private void validateFilter(JsonNode raw, Set<String> fieldIds) {
        if (!raw.isObject()) throw ServiceException.validation("Pivot filter must be an object");
        ObjectNode filter = (ObjectNode) raw;
        String kind = SnapshotMutationSupport.text(filter, "kind");
        if (!FILTER_KINDS.contains(kind)) throw ServiceException.validation("Pivot filter kind is invalid");
        switch (kind) {
            case "manual" -> {
                SnapshotMutationSupport.validateKnownKeys(filter, Set.of("kind", "fieldId", "mode", "memberKeys"), "Pivot manual filter");
                requireField(fieldIds, filter);
                String mode = SnapshotMutationSupport.text(filter, "mode");
                if (!MANUAL_MODES.contains(mode)) throw ServiceException.validation("Pivot manual filter mode is invalid");
                validateMemberValues(filter.get("memberKeys"), "Pivot manual filter memberKeys");
                if ("all".equals(mode) && filter.path("memberKeys").size() != 0) throw ServiceException.validation("Pivot manual filter all mode cannot contain memberKeys");
            }
            case "condition" -> {
                SnapshotMutationSupport.validateKnownKeys(filter, Set.of("kind", "fieldId", "operator", "value"), "Pivot condition filter");
                requireField(fieldIds, filter);
                if (!Set.of("equals", "not-equals", "contains", "greater-than", "greater-or-equal", "less-than", "less-or-equal").contains(SnapshotMutationSupport.text(filter, "operator"))) {
                    throw ServiceException.validation("Pivot condition filter operator is invalid");
                }
                if (!isScalar(filter.get("value"))) throw ServiceException.validation("Pivot condition filter value is invalid");
            }
            case "top-items" -> {
                SnapshotMutationSupport.validateKnownKeys(filter, Set.of("kind", "fieldId", "count", "valueFieldId", "direction"), "Pivot top-items filter");
                requireField(fieldIds, filter);
                JsonNode count = filter.get("count");
                if (count == null || !count.isIntegralNumber() || count.intValue() < 1) throw ServiceException.validation("Pivot top-items count is invalid");
                if (!Set.of("top", "bottom").contains(SnapshotMutationSupport.text(filter, "direction"))) throw ServiceException.validation("Pivot top-items direction is invalid");
                requireField(fieldIds, filter, "valueFieldId");
            }
            default -> throw ServiceException.validation("Pivot filter kind is invalid");
        }
    }

    private void validateValue(JsonNode raw, Set<String> fieldIds) {
        if (!raw.isObject()) throw ServiceException.validation("Pivot value field must be an object");
        ObjectNode value = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(value, VALUE_KEYS, "Pivot value field");
        requireField(fieldIds, value);
        if (!AGGREGATORS.contains(SnapshotMutationSupport.text(value, "summarizeBy"))) throw ServiceException.validation("Pivot value aggregator is invalid");
        if (value.has("baseFieldId")) requireField(fieldIds, value, "baseFieldId");
        if (value.has("baseItem") && !isScalar(value.get("baseItem")) && !value.get("baseItem").isObject()) throw ServiceException.validation("Pivot baseItem is invalid");
        validateShowAs(value.get("showAs"));
    }

    private void validateCalculated(JsonNode raw, Set<String> fieldIds, boolean item) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isArray() || raw.size() > 10_000) throw ServiceException.validation("Pivot calculated definitions are invalid");
        for (JsonNode value : raw) {
            if (!value.isObject()) throw ServiceException.validation("Pivot calculated definition must be an object");
            ObjectNode definition = (ObjectNode) value;
            SnapshotMutationSupport.validateKnownKeys(definition, item ? Set.of("fieldId", "targetFieldId", "name", "formula") : Set.of("fieldId", "name", "formula"), item ? "Pivot calculated item" : "Pivot calculated field");
            requireField(fieldIds, definition);
            if (item) requireField(fieldIds, definition, "targetFieldId");
            SnapshotMutationSupport.text(definition, "name");
            SnapshotMutationSupport.text(definition, "formula");
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

    private void validateSort(JsonNode raw, Set<String> fieldIds, String label) {
        if (raw == null || raw.isNull()) return;
        if (!raw.isObject()) throw ServiceException.validation(label + " must be an object");
        ObjectNode sort = (ObjectNode) raw;
        SnapshotMutationSupport.validateKnownKeys(sort, Set.of("direction", "by", "valueFieldId"), label);
        if (!Set.of("ascending", "descending").contains(SnapshotMutationSupport.text(sort, "direction"))) throw ServiceException.validation(label + " direction is invalid");
        if (sort.has("by") && !Set.of("label", "value").contains(SnapshotMutationSupport.text(sort, "by"))) throw ServiceException.validation(label + " by is invalid");
        if (sort.has("valueFieldId")) requireField(fieldIds, sort, "valueFieldId");
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

    private void validateMemberValues(JsonNode raw, String label) {
        if (raw == null || !raw.isArray() || raw.size() > 10_000) throw ServiceException.validation(label + " must be an array");
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
        if (raw == null || !raw.isArray() || raw.size() > 10_000) throw ServiceException.validation(label + " must be an array");
        for (JsonNode value : raw) if (!isScalar(value)) throw ServiceException.validation(label + " entries must be scalar values");
    }

    private void requireField(Set<String> fieldIds, ObjectNode object) {
        requireField(fieldIds, object, "fieldId");
    }

    private void requireField(Set<String> fieldIds, ObjectNode object, String property) {
        JsonNode value = object.get(property);
        if (value == null || !value.isTextual() || value.asText().isBlank() || !fieldIds.contains(value.asText())) throw ServiceException.validation("Pivot " + property + " is unknown");
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
        String kind = SnapshotMutationSupport.text(source, "kind");
        return switch (kind) {
            case "worksheet-range" -> List.of(SnapshotMutationSupport.range(root, source.get("range")));
            case "worksheet-ranges" -> {
                List<RangeRef> ranges = new ArrayList<>();
                for (JsonNode raw : SnapshotMutationSupport.requiredArray(source, "ranges")) ranges.add(SnapshotMutationSupport.range(root, raw));
                yield List.copyOf(ranges);
            }
            case "table" -> resolveTableRanges(root, SnapshotMutationSupport.text(source, "tableId"));
            case "data-source" -> resolveDataSourceRanges(root, SnapshotMutationSupport.text(source, "dataSourceId"));
            case "named-range" -> List.of();
            default -> throw ServiceException.validation("Pivot source kind is invalid");
        };
    }

    static void forEachWorksheetSourceRange(ObjectNode pivot, Consumer<JsonNode> consumer) {
        JsonNode sourceRaw = pivot.get("source");
        if (sourceRaw == null || !sourceRaw.isObject()) return;
        ObjectNode source = (ObjectNode) sourceRaw;
        switch (source.path("kind").asText()) {
            case "worksheet-range" -> consumer.accept(source.get("range"));
            case "worksheet-ranges" -> {
                for (JsonNode range : source.path("ranges")) consumer.accept(range);
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

    private static List<RangeRef> resolveTableRanges(ObjectNode root, String tableId) {
        List<RangeRef> ranges = new ArrayList<>();
        JsonNode workbookTables = root.path("dataModel").get("tables");
        if (workbookTables != null && workbookTables.isArray()) {
            for (JsonNode raw : workbookTables) if (raw.isObject() && tableId.equals(raw.path("id").asText())) {
                JsonNode range = raw.get("sourceRange");
                if (range != null) ranges.add(SnapshotMutationSupport.range(root, range));
            }
        }
        for (JsonNode rawSheet : SnapshotMutationSupport.sheets(root)) {
            if (!rawSheet.isObject()) continue;
            ObjectNode sheet = (ObjectNode) rawSheet;
            JsonNode tables = sheet.get("sheetTables");
            if (tables == null || !tables.isArray()) continue;
            for (JsonNode raw : tables) if (raw.isObject() && tableId.equals(raw.path("id").asText())) {
                JsonNode range = raw.get("range");
                if (range != null) ranges.add(SnapshotMutationSupport.range(root, range));
            }
        }
        if (ranges.isEmpty()) throw ServiceException.notFound("Pivot table source not found: " + tableId);
        return List.copyOf(ranges);
    }

    private static List<RangeRef> resolveDataSourceRanges(ObjectNode root, String sourceId) {
        JsonNode rawSources = root.path("dataModel").get("sources");
        if (rawSources == null || !rawSources.isArray()) return List.of();
        for (JsonNode raw : rawSources) if (raw.isObject() && sourceId.equals(raw.path("id").asText())) {
            ObjectNode source = (ObjectNode) raw;
            List<RangeRef> ranges = new ArrayList<>();
            JsonNode single = source.get("sourceRange");
            if (single != null && !single.isNull()) ranges.add(SnapshotMutationSupport.range(root, single));
            JsonNode many = source.get("ranges");
            if (many != null && many.isArray()) for (JsonNode range : many) ranges.add(SnapshotMutationSupport.range(root, range));
            return List.copyOf(ranges);
        }
        return List.of();
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
