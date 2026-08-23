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

/** Reducers for persisted pivot definitions; result trees remain derived. */
final class PivotMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of("pivot.add", "pivot.remove", "pivot.update", "pivot.refresh");

    PivotMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, true, "structure");
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported pivot mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        return switch (id()) {
            case "pivot.add" -> pivotRanges(root, mutation.sheetId(), SnapshotMutationSupport.params(mutation));
            case "pivot.remove" -> pivotRanges(root, mutation.sheetId(), currentPivot(root, mutation.sheetId(), pivotId(mutation.params())));
            case "pivot.update", "pivot.refresh" -> pivotRanges(root, mutation.sheetId(), currentPivot(root, mutation.sheetId(), pivotId(SnapshotMutationSupport.params(mutation))));
            default -> throw ServiceException.validation("Unsupported pivot mutation: " + id());
        };
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        ArrayNode pivots = SnapshotMutationSupport.array(sheet, "pivots");
        switch (id()) {
            case "pivot.add" -> add(root, mutation.sheetId(), pivots, SnapshotMutationSupport.params(mutation));
            case "pivot.remove" -> remove(pivots, pivotId(mutation.params()));
            case "pivot.update" -> update(root, mutation.sheetId(), pivots, SnapshotMutationSupport.params(mutation));
            case "pivot.refresh" -> refresh(pivots, SnapshotMutationSupport.params(mutation));
            default -> throw ServiceException.validation("Unsupported pivot mutation: " + id());
        }
        return root;
    }

    private void add(ObjectNode root, String sheetId, ArrayNode pivots, ObjectNode pivot) {
        validatePivot(root, sheetId, pivot);
        String pivotId = SnapshotMutationSupport.text(pivot, "id");
        for (JsonNode sheet : SnapshotMutationSupport.sheets(root)) {
            if (!sheet.isObject()) continue;
            if (SnapshotMutationSupport.findById(SnapshotMutationSupport.array((ObjectNode) sheet, "pivots"), pivotId) != null) {
                throw ServiceException.conflict("Pivot already exists: " + pivotId);
            }
        }
        pivots.add(pivot.deepCopy());
    }

    private void remove(ArrayNode pivots, String pivotId) {
        if (!SnapshotMutationSupport.removeById(pivots, pivotId)) throw ServiceException.notFound("Pivot not found: " + pivotId);
    }

    private void update(ObjectNode root, String sheetId, ArrayNode pivots, ObjectNode params) {
        ObjectNode pivot = SnapshotMutationSupport.requireById(pivots, pivotId(params), "Pivot");
        ObjectNode next = pivot.deepCopy();
        for (String property : List.of("sourceRange", "dataSource", "layout", "slicers", "timelines", "chartReferences")) {
            JsonNode value = params.get(property);
            if (value != null) next.set(property, value.deepCopy());
        }
        validatePivot(root, sheetId, next);
        pivots.set(SnapshotMutationSupport.indexById(pivots, pivotId(params)), next);
    }

    private void refresh(ArrayNode pivots, ObjectNode params) {
        ObjectNode pivot = SnapshotMutationSupport.requireById(pivots, pivotId(params), "Pivot");
        JsonNode revision = params.get("refreshRevision");
        JsonNode refreshedAt = params.get("lastRefreshedAt");
        if (refreshedAt == null || !refreshedAt.isTextual()) throw ServiceException.validation("Pivot lastRefreshedAt must be a string");
        if (revision == null || !revision.isIntegralNumber() || revision.longValue() < 0) throw ServiceException.validation("Pivot refresh revision is invalid");
        pivot.set("refreshRevision", revision.deepCopy());
        if (refreshedAt.asText().isBlank()) pivot.remove("lastRefreshedAt");
        else pivot.set("lastRefreshedAt", refreshedAt.deepCopy());
    }

    private ObjectNode currentPivot(ObjectNode root, String sheetId, String pivotId) {
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, sheetId);
        return SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(sheet, "pivots"), pivotId, "Pivot");
    }

    private String pivotId(ObjectNode params) {
        return SnapshotMutationSupport.text(params, "pivotId");
    }

    private String pivotId(JsonNode value) {
        if (value == null || !value.isTextual() || value.asText().isBlank()) throw ServiceException.validation("Pivot id is required");
        return value.asText();
    }

    private void validatePivot(ObjectNode root, String expectedSheetId, ObjectNode pivot) {
        SnapshotMutationSupport.text(pivot, "id");
        SnapshotMutationSupport.requireEntitySheet(pivot, expectedSheetId);
        sourceRange(root, pivot.get("sourceRange"));
        JsonNode layout = pivot.get("layout");
        if (layout == null || !layout.isObject()) throw ServiceException.validation("Pivot layout is required");
        for (String property : List.of("rows", "columns", "filters", "values")) {
            if (!layout.path(property).isArray()) throw ServiceException.validation("Pivot layout " + property + " must be an array");
        }
        if (!layout.path("showSubtotals").isBoolean() || !layout.path("showGrandTotals").isBoolean()) throw ServiceException.validation("Pivot layout totals flags are required");
        JsonNode source = pivot.get("dataSource");
        if (source != null && !source.isNull()) validateDataSource(root, source);
        validatePivotArray(pivot.get("slicers"), "Pivot slicers");
        validatePivotArray(pivot.get("timelines"), "Pivot timelines");
        validatePivotArray(pivot.get("chartReferences"), "Pivot chartReferences");
    }

    private void validateDataSource(ObjectNode root, JsonNode source) {
        if (!source.isObject()) throw ServiceException.validation("Pivot data source must be an object");
        String kind = source.path("kind").asText();
        if (kind.equals("worksheet-range")) {
            sourceRange(root, source.get("range"));
            return;
        }
        if (kind.equals("worksheet-ranges")) {
            JsonNode ranges = source.get("ranges");
            if (ranges == null || !ranges.isArray() || ranges.isEmpty()) throw ServiceException.validation("Pivot data source ranges are required");
            for (JsonNode range : ranges) sourceRange(root, range);
            if (!source.path("relationships").isArray()) throw ServiceException.validation("Pivot data source relationships must be an array");
            return;
        }
        throw ServiceException.validation("Pivot data source kind is invalid");
    }

    private void validatePivotArray(JsonNode value, String name) {
        if (value != null && (!value.isArray() || value.size() > 10_000)) throw ServiceException.validation(name + " is invalid");
    }

    private RangeRef sourceRange(ObjectNode root, JsonNode value) {
        return SnapshotMutationSupport.range(root, value);
    }

    private List<RangeRef> pivotRanges(ObjectNode root, String sheetId, ObjectNode pivot) {
        List<RangeRef> ranges = new ArrayList<>();
        ranges.add(SnapshotMutationSupport.wholeSheetRange(root, sheetId));
        ranges.add(sourceRange(root, pivot.get("sourceRange")));
        JsonNode source = pivot.get("dataSource");
        if (source != null && source.isObject() && "worksheet-ranges".equals(source.path("kind").asText())) {
            for (JsonNode range : source.path("ranges")) ranges.add(sourceRange(root, range));
        } else if (source != null && source.isObject() && "worksheet-range".equals(source.path("kind").asText())) {
            ranges.add(sourceRange(root, source.get("range")));
        }
        return ranges.stream().distinct().toList();
    }
}
