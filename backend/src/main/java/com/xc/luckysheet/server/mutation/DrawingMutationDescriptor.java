package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;

import java.util.List;
import java.util.Set;

/** Reducers for the canonical DrawingObject and DrawingPayload collections. */
final class DrawingMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "drawing.add", "drawing.remove", "drawing.transform", "drawing.transform.batch",
            "drawing.anchor", "drawing.payload.update", "drawing.zorder", "drawing.zorder.restore"
    );

    DrawingMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR, true, "drawing");
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported drawing mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        SnapshotMutationSupport.params(mutation);
        return List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        switch (id()) {
            case "drawing.add" -> add(root, sheet, mutation.sheetId(), params);
            case "drawing.remove" -> remove(sheet, params);
            case "drawing.transform" -> transform(sheet, params);
            case "drawing.transform.batch" -> transformBatch(sheet, params);
            case "drawing.anchor" -> anchor(root, sheet, mutation.sheetId(), params);
            case "drawing.payload.update" -> payload(sheet, params);
            case "drawing.zorder" -> zOrder(sheet, params);
            case "drawing.zorder.restore" -> restoreZOrder(sheet, params);
            default -> throw ServiceException.validation("Unsupported drawing mutation: " + id());
        }
        return root;
    }

    private void add(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        ObjectNode drawing = SnapshotMutationSupport.requiredObject(params, "drawing");
        ObjectNode payload = SnapshotMutationSupport.requiredObject(params, "payload");
        SnapshotMutationSupport.requireEntitySheet(params, sheetId);
        validateDrawing(root, drawing, payload, sheetId);
        ArrayNode drawings = SnapshotMutationSupport.array(sheet, "drawings");
        ObjectNode payloads = SnapshotMutationSupport.object(sheet, "drawingPayloads");
        String drawingId = SnapshotMutationSupport.text(drawing, "id");
        String payloadId = SnapshotMutationSupport.text(drawing, "payloadId");
        if (SnapshotMutationSupport.findById(drawings, drawingId) != null) throw ServiceException.conflict("Drawing already exists: " + drawingId);
        if (payloads.has(payloadId)) throw ServiceException.conflict("Drawing payload already exists: " + payloadId);
        drawings.add(drawing.deepCopy());
        payloads.set(payloadId, payload.deepCopy());
    }

    private void remove(ObjectNode sheet, ObjectNode params) {
        String drawingId = SnapshotMutationSupport.text(params, "drawingId");
        ArrayNode drawings = SnapshotMutationSupport.array(sheet, "drawings");
        ObjectNode drawing = SnapshotMutationSupport.requireById(drawings, drawingId, "Drawing");
        String payloadId = SnapshotMutationSupport.text(drawing, "payloadId");
        ObjectNode payloads = SnapshotMutationSupport.object(sheet, "drawingPayloads");
        if (!payloads.has(payloadId)) throw ServiceException.validation("Drawing payload is missing: " + payloadId);
        SnapshotMutationSupport.removeById(drawings, drawingId);
        payloads.remove(payloadId);
    }

    private void transform(ObjectNode sheet, ObjectNode params) {
        ObjectNode drawing = drawing(sheet, SnapshotMutationSupport.text(params, "drawingId"));
        ObjectNode transform = SnapshotMutationSupport.requiredObject(params, "transform");
        validateTransform(transform);
        drawing.set("transform", transform.deepCopy());
    }

    private void transformBatch(ObjectNode sheet, ObjectNode params) {
        ArrayNode entries = SnapshotMutationSupport.requiredArray(params, "entries");
        if (entries.isEmpty() || entries.size() > 1_000) throw ServiceException.validation("Drawing transform batch is invalid");
        for (JsonNode raw : entries) {
            if (!raw.isObject()) throw ServiceException.validation("Drawing transform entry must be an object");
            ObjectNode entry = (ObjectNode) raw;
            ObjectNode drawing = drawing(sheet, SnapshotMutationSupport.text(entry, "drawingId"));
            ObjectNode before = SnapshotMutationSupport.requiredObject(entry, "before");
            ObjectNode after = SnapshotMutationSupport.requiredObject(entry, "after");
            validateTransform(before);
            validateTransform(after);
            if (!drawing.path("transform").equals(before)) throw ServiceException.conflict("Drawing changed before transform batch: " + drawing.path("id").asText());
        }
        for (JsonNode raw : entries) {
            ObjectNode entry = (ObjectNode) raw;
            drawing(sheet, SnapshotMutationSupport.text(entry, "drawingId")).set("transform", entry.get("after").deepCopy());
        }
    }

    private void anchor(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        ObjectNode drawing = drawing(sheet, SnapshotMutationSupport.text(params, "drawingId"));
        ObjectNode anchor = SnapshotMutationSupport.requiredObject(params, "anchor");
        validateAnchor(root, sheetId, anchor);
        drawing.set("anchor", anchor.deepCopy());
    }

    private void payload(ObjectNode sheet, ObjectNode params) {
        String payloadId = SnapshotMutationSupport.text(params, "payloadId");
        ObjectNode before = SnapshotMutationSupport.requiredObject(params, "before");
        ObjectNode after = SnapshotMutationSupport.requiredObject(params, "after");
        ObjectNode drawing = drawingByPayload(sheet, payloadId);
        validatePayloadPair(drawing, after);
        ObjectNode payloads = SnapshotMutationSupport.object(sheet, "drawingPayloads");
        JsonNode current = payloads.get(payloadId);
        if (current == null || !current.isObject()) throw ServiceException.notFound("Drawing payload not found: " + payloadId);
        if (!current.equals(before)) throw ServiceException.conflict("Drawing payload changed before update: " + payloadId);
        payloads.set(payloadId, after.deepCopy());
    }

    private void zOrder(ObjectNode sheet, ObjectNode params) {
        ObjectNode drawing = drawing(sheet, SnapshotMutationSupport.text(params, "drawingId"));
        String direction = SnapshotMutationSupport.text(params, "direction");
        if (!Set.of("forward", "backward", "front", "back").contains(direction)) throw ServiceException.validation("Drawing z-order direction is invalid");
        ArrayNode drawings = SnapshotMutationSupport.array(sheet, "drawings");
        List<ObjectNode> ordered = drawings.findValuesAsText("id").stream().map(id -> SnapshotMutationSupport.requireById(drawings, id, "Drawing"))
                .sorted((left, right) -> Double.compare(left.path("zIndex").asDouble(), right.path("zIndex").asDouble()))
                .toList();
        int current = ordered.indexOf(drawing);
        if (current < 0) throw ServiceException.notFound("Drawing not found");
        if (direction.equals("front")) {
            double max = ordered.stream().mapToDouble(item -> item.path("zIndex").asDouble()).max().orElse(0);
            drawing.put("zIndex", max + 1);
            return;
        }
        if (direction.equals("back")) {
            double min = ordered.stream().mapToDouble(item -> item.path("zIndex").asDouble()).min().orElse(0);
            drawing.put("zIndex", min - 1);
            return;
        }
        int target = direction.equals("forward") ? current + 1 : current - 1;
        if (target < 0 || target >= ordered.size()) return;
        ObjectNode peer = ordered.get(target);
        JsonNode original = drawing.get("zIndex");
        drawing.set("zIndex", peer.get("zIndex").deepCopy());
        peer.set("zIndex", original.deepCopy());
    }

    private void restoreZOrder(ObjectNode sheet, ObjectNode params) {
        ArrayNode entries = SnapshotMutationSupport.requiredArray(params, "entries");
        if (entries.size() > 10_000) throw ServiceException.validation("Drawing z-order restore is too large");
        for (JsonNode raw : entries) {
            if (!raw.isObject()) throw ServiceException.validation("Drawing z-order entry must be an object");
            ObjectNode entry = (ObjectNode) raw;
            JsonNode zIndex = entry.get("zIndex");
            if (zIndex == null || !zIndex.isNumber() || !Double.isFinite(zIndex.asDouble())) throw ServiceException.validation("Drawing zIndex is invalid");
            drawing(sheet, SnapshotMutationSupport.text(entry, "drawingId")).set("zIndex", zIndex.deepCopy());
        }
    }

    private ObjectNode drawing(ObjectNode sheet, String id) {
        return SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(sheet, "drawings"), id, "Drawing");
    }

    private ObjectNode drawingByPayload(ObjectNode sheet, String payloadId) {
        for (JsonNode raw : SnapshotMutationSupport.array(sheet, "drawings")) {
            if (raw.isObject() && payloadId.equals(raw.path("payloadId").asText())) return (ObjectNode) raw;
        }
        throw ServiceException.notFound("Drawing payload not found: " + payloadId);
    }

    private void validateDrawing(ObjectNode root, ObjectNode drawing, ObjectNode payload, String sheetId) {
        SnapshotMutationSupport.text(drawing, "id");
        SnapshotMutationSupport.requireEntitySheet(drawing, sheetId);
        SnapshotMutationSupport.text(drawing, "payloadId");
        validateAnchor(root, sheetId, SnapshotMutationSupport.requiredObject(drawing, "anchor"));
        validateTransformObject(drawing.path("transform"));
        JsonNode zIndex = drawing.get("zIndex");
        if (zIndex == null || !zIndex.isNumber() || !Double.isFinite(zIndex.asDouble())) throw ServiceException.validation("Drawing zIndex is invalid");
        validatePayloadPair(drawing, payload);
    }

    private void validatePayloadPair(ObjectNode drawing, ObjectNode payload) {
        String drawingKind = SnapshotMutationSupport.text(drawing, "kind");
        String payloadKind = SnapshotMutationSupport.text(payload, "kind");
        if (!Set.of("image", "shape", "textbox", "chart").contains(drawingKind) || !drawingKind.equals(payloadKind)) {
            throw ServiceException.validation("Drawing and payload kinds must match");
        }
        if (payloadKind.equals("chart") && !SnapshotMutationSupport.text(drawing, "payloadId").equals(SnapshotMutationSupport.text(payload, "chartId"))) {
            throw ServiceException.validation("Chart payload identity does not match drawing payloadId");
        }
    }

    private void validateAnchor(ObjectNode root, String sheetId, ObjectNode anchor) {
        validateAnchorWithoutBounds(anchor);
        String kind = SnapshotMutationSupport.text(anchor, "kind");
        if (kind.equals("absolute")) return;
        ObjectNode start = JsonNodeFactoryHolder.coordinate(anchor.get("row"), anchor.get("column"));
        SnapshotMutationSupport.coordinate(root, sheetId, start);
        if (kind.equals("two-cell")) {
            ObjectNode end = JsonNodeFactoryHolder.coordinate(anchor.get("endRow"), anchor.get("endColumn"));
            SnapshotMutationSupport.CellCoordinate startCoordinate = SnapshotMutationSupport.coordinate(root, sheetId, start);
            SnapshotMutationSupport.CellCoordinate endCoordinate = SnapshotMutationSupport.coordinate(root, sheetId, end);
            if (endCoordinate.row() < startCoordinate.row() || endCoordinate.column() < startCoordinate.column()) throw ServiceException.validation("Two-cell drawing anchor bounds are invalid");
        }
    }

    private void validateAnchorWithoutBounds(JsonNode raw) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Drawing anchor must be an object");
        ObjectNode anchor = (ObjectNode) raw;
        String kind = SnapshotMutationSupport.text(anchor, "kind");
        if (!Set.of("absolute", "one-cell", "two-cell").contains(kind)) throw ServiceException.validation("Drawing anchor kind is invalid");
        if (kind.equals("absolute")) return;
        integer(anchor.get("row"), "Drawing anchor row");
        integer(anchor.get("column"), "Drawing anchor column");
        if (kind.equals("two-cell")) {
            integer(anchor.get("endRow"), "Drawing anchor endRow");
            integer(anchor.get("endColumn"), "Drawing anchor endColumn");
        }
    }

    private void validateTransform(ObjectNode transform) {
        validateTransformObject(transform);
    }

    private void validateTransformObject(JsonNode raw) {
        if (raw == null || !raw.isObject()) throw ServiceException.validation("Drawing transform must be an object");
        ObjectNode transform = (ObjectNode) raw;
        for (String key : List.of("x", "y", "width", "height")) {
            JsonNode value = transform.get(key);
            if (value == null || !value.isNumber() || !Double.isFinite(value.asDouble())) throw ServiceException.validation("Drawing transform " + key + " is invalid");
        }
        if (transform.path("width").asDouble() < 0 || transform.path("height").asDouble() < 0) throw ServiceException.validation("Drawing transform dimensions are invalid");
        JsonNode rotation = transform.get("rotation");
        if (rotation != null && (!rotation.isNumber() || !Double.isFinite(rotation.asDouble()))) throw ServiceException.validation("Drawing rotation is invalid");
    }

    private void integer(JsonNode value, String name) {
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 0) throw ServiceException.validation(name + " is invalid");
    }

    private static final class JsonNodeFactoryHolder {
        private JsonNodeFactoryHolder() {
        }

        static ObjectNode coordinate(JsonNode row, JsonNode column) {
            ObjectNode coordinate = com.fasterxml.jackson.databind.node.JsonNodeFactory.instance.objectNode();
            coordinate.set("row", row);
            coordinate.set("column", column);
            return coordinate;
        }
    }
}
