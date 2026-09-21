package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

/** Server reducer for the persisted shared dashboard/analysis view state. */
final class AnalysisViewMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of("analysis.view.replace");
    private static final Set<String> VIEW_KEYS = Set.of(
            "kind", "id", "name", "tableId", "fields", "groupBy", "sort", "filters", "charts", "layout", "revision"
    );
    private static final Set<String> FIELD_KEYS = Set.of("fieldId", "caption", "formula", "widthPx");
    private static final Set<String> SORT_KEYS = Set.of("fieldId", "direction");
    private static final Set<String> FILTER_KEYS = Set.of("id", "fieldId", "operator", "values");
    private static final Set<String> CHART_KEYS = Set.of("chartId", "fieldMap");
    private static final Set<String> FIELD_MAP_KEYS = Set.of("category", "series", "value", "color", "size", "tooltip");
    private static final Set<String> LAYOUT_KEYS = Set.of("columns", "rowHeightPx", "gapPx");
    private static final Set<String> FILTER_OPERATORS = Set.of("equals", "not-equals", "contains", "in", "between");

    AnalysisViewMutationDescriptor() {
        super("analysis.view.replace", WorkbookAclRole.EDITOR);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        validate(snapshot, mutation);
        return List.of();
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        validate(root, mutation, params);
        ArrayNode views = SnapshotMutationSupport.dataModelArray(root, "views");
        JsonNode rawView = params.get("view");
        if (rawView == null || rawView.isNull()) {
            String viewId = SnapshotMutationSupport.text(params, "viewId");
            ObjectNode current = SnapshotMutationSupport.requireById(views, viewId, "Analysis view");
            if (!"analysis".equals(current.path("kind").asText())) throw ServiceException.conflict("A table view owns this id: " + viewId);
            views.remove(SnapshotMutationSupport.indexById(views, viewId));
            return root;
        }
        ObjectNode view = (ObjectNode) rawView;
        String viewId = SnapshotMutationSupport.text(view, "id");
        int index = SnapshotMutationSupport.indexById(views, viewId);
        if (index >= 0 && !"analysis".equals(views.get(index).path("kind").asText())) {
            throw ServiceException.conflict("A table view owns this id: " + viewId);
        }
        if (index >= 0) views.set(index, view.deepCopy());
        else views.add(view.deepCopy());
        return root;
    }

    private void validate(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        validate(root, mutation, SnapshotMutationSupport.params(mutation));
    }

    private void validate(ObjectNode root, OperationMutation mutation, ObjectNode params) {
        SnapshotMutationSupport.validateKnownKeys(params, Set.of("view", "viewId"), "analysis.view.replace");
        JsonNode rawView = params.get("view");
        if (rawView == null) throw ServiceException.validation("analysis.view.replace requires view");
        if (rawView.isNull()) {
            SnapshotMutationSupport.text(params, "viewId");
            return;
        }
        if (!rawView.isObject()) throw ServiceException.validation("Analysis view must be an object or null");
        validateView(root, mutation, (ObjectNode) rawView);
    }

    private void validateView(ObjectNode root, OperationMutation mutation, ObjectNode view) {
        SnapshotMutationSupport.validateKnownKeys(view, VIEW_KEYS, "Analysis view");
        if (!"analysis".equals(SnapshotMutationSupport.text(view, "kind"))) throw ServiceException.validation("Analysis view kind is invalid");
        String viewId = SnapshotMutationSupport.text(view, "id");
        if (viewId.length() > 200) throw ServiceException.validation("Analysis view id is too long");
        if (SnapshotMutationSupport.text(view, "name").length() > 255) throw ServiceException.validation("Analysis view name is too long");
        String tableId = SnapshotMutationSupport.text(view, "tableId");
        ObjectNode table = SnapshotMutationSupport.requireById(SnapshotMutationSupport.dataModelArray(root, "tables"), tableId, "Workbook table");
        Set<String> fieldIds = tableFieldIds(table);

        ArrayNode fields = SnapshotMutationSupport.requiredArray(view, "fields");
        if (fields.size() > 16_384) throw ServiceException.validation("Analysis view has too many fields");
        Set<String> selected = new HashSet<>();
        for (JsonNode raw : fields) {
            if (!raw.isObject()) throw ServiceException.validation("Analysis view field must be an object");
            ObjectNode field = (ObjectNode) raw;
            SnapshotMutationSupport.validateKnownKeys(field, FIELD_KEYS, "Analysis view field");
            String fieldId = SnapshotMutationSupport.text(field, "fieldId");
            SnapshotMutationSupport.text(field, "caption");
            if (!fieldIds.contains(fieldId)) throw ServiceException.validation("Analysis view field is not present in the table: " + fieldId);
            if (!selected.add(fieldId)) throw ServiceException.validation("Analysis view field is duplicated: " + fieldId);
            if (field.has("widthPx") && (!field.get("widthPx").isNumber() || !Double.isFinite(field.get("widthPx").asDouble()) || field.get("widthPx").asDouble() <= 0)) {
                throw ServiceException.validation("Analysis view field widthPx is invalid");
            }
        }
        validateOptionalFieldList(view.get("groupBy"), fieldIds, "Analysis view groupBy");
        validateSort(view.get("sort"), fieldIds);
        validateFilters(view.get("filters"), fieldIds);
        validateCharts(view.get("charts"), fieldIds);
        validateLayout(view.get("layout"));
        JsonNode revision = view.get("revision");
        if (revision == null || !revision.isIntegralNumber() || revision.asLong() < 0) throw ServiceException.validation("Analysis view revision is invalid");
        if (mutation.sheetId().isBlank()) throw ServiceException.validation("Analysis view mutation sheetId is required");
    }

    private Set<String> tableFieldIds(ObjectNode table) {
        ArrayNode fields = SnapshotMutationSupport.requiredArray(table, "fields");
        Set<String> ids = new HashSet<>();
        for (JsonNode raw : fields) {
            if (!raw.isObject()) throw ServiceException.validation("Workbook table field must be an object");
            String id = SnapshotMutationSupport.text((ObjectNode) raw, "id");
            if (!ids.add(id)) throw ServiceException.validation("Workbook table field is duplicated: " + id);
        }
        return ids;
    }

    private void validateOptionalFieldList(JsonNode value, Set<String> fieldIds, String label) {
        if (value == null || value.isNull()) return;
        if (!value.isArray()) throw ServiceException.validation(label + " must be an array");
        for (JsonNode raw : value) {
            if (!raw.isTextual() || !fieldIds.contains(raw.asText())) throw ServiceException.validation(label + " contains an unknown field");
        }
    }

    private void validateSort(JsonNode value, Set<String> fieldIds) {
        if (value == null || value.isNull()) return;
        if (!value.isArray()) throw ServiceException.validation("Analysis view sort must be an array");
        for (JsonNode raw : value) {
            if (!raw.isObject()) throw ServiceException.validation("Analysis view sort entry must be an object");
            ObjectNode sort = (ObjectNode) raw;
            SnapshotMutationSupport.validateKnownKeys(sort, SORT_KEYS, "Analysis view sort entry");
            if (!fieldIds.contains(SnapshotMutationSupport.text(sort, "fieldId")) || !Set.of("asc", "desc").contains(SnapshotMutationSupport.text(sort, "direction"))) {
                throw ServiceException.validation("Analysis view sort entry is invalid");
            }
        }
    }

    private void validateFilters(JsonNode value, Set<String> fieldIds) {
        if (value == null || !value.isArray()) throw ServiceException.validation("Analysis view filters are required");
        if (value.size() > 128) throw ServiceException.validation("Analysis view has too many filters");
        Set<String> ids = new HashSet<>();
        for (JsonNode raw : value) {
            if (!raw.isObject()) throw ServiceException.validation("Analysis filter must be an object");
            ObjectNode filter = (ObjectNode) raw;
            SnapshotMutationSupport.validateKnownKeys(filter, FILTER_KEYS, "Analysis filter");
            String id = SnapshotMutationSupport.text(filter, "id");
            if (!ids.add(id) || !fieldIds.contains(SnapshotMutationSupport.text(filter, "fieldId")) || !FILTER_OPERATORS.contains(SnapshotMutationSupport.text(filter, "operator"))) {
                throw ServiceException.validation("Analysis filter identity or field is invalid");
            }
            ArrayNode values = SnapshotMutationSupport.requiredArray(filter, "values");
            if (values.isEmpty() || values.size() > 10_000) throw ServiceException.validation("Analysis filter values are invalid");
            for (JsonNode item : values) if (!isScalar(item)) throw ServiceException.validation("Analysis filter values must be scalar");
        }
    }

    private void validateCharts(JsonNode value, Set<String> fieldIds) {
        if (value == null || !value.isArray()) throw ServiceException.validation("Analysis view charts are required");
        if (value.size() > 100) throw ServiceException.validation("Analysis view has too many chart bindings");
        Set<String> ids = new HashSet<>();
        for (JsonNode raw : value) {
            if (!raw.isObject()) throw ServiceException.validation("Analysis chart binding must be an object");
            ObjectNode chart = (ObjectNode) raw;
            SnapshotMutationSupport.validateKnownKeys(chart, CHART_KEYS, "Analysis chart binding");
            String chartId = SnapshotMutationSupport.text(chart, "chartId");
            if (!ids.add(chartId)) throw ServiceException.validation("Analysis chart binding is duplicated: " + chartId);
            ObjectNode map = SnapshotMutationSupport.requiredObject(chart, "fieldMap");
            SnapshotMutationSupport.validateKnownKeys(map, FIELD_MAP_KEYS, "Analysis chart field map");
            map.fieldNames().forEachRemaining(key -> {
                String fieldId = SnapshotMutationSupport.text(map, key);
                if (!fieldIds.contains(fieldId)) throw ServiceException.validation("Analysis chart field is not present in the table: " + fieldId);
            });
        }
    }

    private void validateLayout(JsonNode value) {
        if (value == null || !value.isObject()) throw ServiceException.validation("Analysis view layout is required");
        ObjectNode layout = (ObjectNode) value;
        SnapshotMutationSupport.validateKnownKeys(layout, LAYOUT_KEYS, "Analysis view layout");
        JsonNode columns = layout.get("columns");
        JsonNode rowHeight = layout.get("rowHeightPx");
        JsonNode gap = layout.get("gapPx");
        if (columns == null || !columns.isIntegralNumber() || columns.asInt() < 1 || columns.asInt() > 12
                || rowHeight == null || !rowHeight.isNumber() || !Double.isFinite(rowHeight.asDouble()) || rowHeight.asDouble() < 1
                || gap == null || !gap.isNumber() || !Double.isFinite(gap.asDouble()) || gap.asDouble() < 0) {
            throw ServiceException.validation("Analysis view layout is invalid");
        }
    }

    private boolean isScalar(JsonNode value) {
        return value.isNull() || value.isTextual() || value.isBoolean()
                || (value.isNumber() && Double.isFinite(value.asDouble()));
    }
}
