package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.service.ServiceException;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

@Component
public class MutationDescriptorRegistry {
    private final Map<String, MutationDescriptor> descriptors = new ConcurrentHashMap<>();

    public MutationDescriptorRegistry() {
        register(new CellSetDescriptor("cell.set", false));
        register(new CellSetDescriptor("cell.restore", false));
        register(new CellSetDescriptor("workbook.restore", true));
    }

    public void register(MutationDescriptor descriptor) {
        if (descriptors.putIfAbsent(descriptor.id(), descriptor) != null) {
            throw new IllegalStateException("Duplicate mutation descriptor: " + descriptor.id());
        }
    }

    public MutationDescriptor require(String id, boolean internalCall) {
        MutationDescriptor descriptor = descriptors.get(id);
        if (descriptor == null) throw ServiceException.validation("Unknown mutation: " + id);
        if (descriptor.internalOnly() && !internalCall) throw ServiceException.forbidden("Mutation is server-only: " + id);
        return descriptor;
    }

    public JsonNode applyPublicMutations(JsonNode snapshot, List<OperationMutation> mutations) {
        JsonNode current = snapshot.deepCopy();
        for (OperationMutation mutation : mutations) {
            current = require(mutation.id(), false).apply(current, mutation);
        }
        return current;
    }

    public List<RangeRef> resolveRanges(JsonNode snapshot, OperationMutation mutation) {
        return require(mutation.id(), false).affectedRanges(snapshot, mutation);
    }

    public List<String> ids() {
        return descriptors.keySet().stream().sorted().toList();
    }

    private static final class CellSetDescriptor implements MutationDescriptor {
        private final String id;
        private final boolean internalOnly;

        private CellSetDescriptor(String id, boolean internalOnly) {
            this.id = id;
            this.internalOnly = internalOnly;
        }

        @Override
        public String id() {
            return id;
        }

        @Override
        public boolean internalOnly() {
            return internalOnly;
        }

        @Override
        public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
            Coordinates coordinates = coordinates(mutation);
            return List.of(new RangeRef(mutation.sheetId(), coordinates.row(), coordinates.row(), coordinates.column(), coordinates.column()));
        }

        @Override
        public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
            if (internalOnly) throw ServiceException.forbidden("Internal mutation cannot be submitted by a client");
            Coordinates coordinates = coordinates(mutation);
            ObjectNode root = requireObject(snapshot, "snapshot");
            JsonNode sheets = root.get("sheets");
            if (sheets == null || !sheets.isArray()) throw ServiceException.validation("Snapshot sheets array is required");
            ObjectNode sheet = null;
            for (JsonNode candidate : sheets) {
                if (candidate.isObject() && mutation.sheetId().equals(candidate.path("id").asText())) {
                    sheet = (ObjectNode) candidate;
                    break;
                }
            }
            if (sheet == null) throw ServiceException.notFound("Sheet not found: " + mutation.sheetId());
            ObjectNode cells = sheet.with("cells");
            ObjectNode row = cells.with(String.valueOf(coordinates.row()));
            if ("cell.restore".equals(id)) {
                JsonNode previous = mutation.params().get("previous");
                if (previous == null || previous.isNull()) row.remove(String.valueOf(coordinates.column()));
                else row.set(String.valueOf(coordinates.column()), previous.deepCopy());
            } else {
                JsonNode value = mutation.params().get("value");
                if (value == null || !value.isObject()) throw ServiceException.validation("cell.set value must be an object");
                row.set(String.valueOf(coordinates.column()), value.deepCopy());
            }
            return root;
        }

        private Coordinates coordinates(OperationMutation mutation) {
            if (mutation.params() == null || !mutation.params().isObject()) throw ServiceException.validation("Mutation params must be an object");
            JsonNode params = mutation.params();
            String declaredSheet = params.path("sheetId").asText("");
            if (!declaredSheet.isBlank() && !declaredSheet.equals(mutation.sheetId())) throw ServiceException.validation("Mutation sheetId does not match params");
            int row = params.path("row").asInt(-1);
            int column = params.path("column").asInt(-1);
            if (row < 0 || column < 0 || row > 1_048_575 || column > 16_383) throw ServiceException.validation("Cell coordinates are out of bounds");
            return new Coordinates(row, column);
        }

        private ObjectNode requireObject(JsonNode node, String name) {
            if (!node.isObject()) throw ServiceException.validation(name + " must be an object");
            return (ObjectNode) node;
        }
    }

    private record Coordinates(int row, int column) {
    }
}
