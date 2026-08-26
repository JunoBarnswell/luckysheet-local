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
import java.util.Comparator;
import java.util.List;
import java.util.Set;

/** Reducers for SparklineModel and its canonical group membership state. */
final class SparklineMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of(
            "sparkline.add", "sparkline.remove", "sparkline.update",
            "sparkline.group.add", "sparkline.group.remove", "sparkline.group.replace"
    );

    SparklineMutationDescriptor(String id) {
        super(id, WorkbookAclRole.EDITOR);
        if (!IDS.contains(id)) throw new IllegalArgumentException("Unsupported sparkline mutation: " + id);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        List<RangeRef> ranges = new ArrayList<>();
        ranges.add(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
        if (id().equals("sparkline.add")) ranges.add(sourceRange(root, SnapshotMutationSupport.requiredObject(params, "sparkline")));
        if (id().equals("sparkline.update")) {
            ObjectNode current = sparkline(SnapshotMutationSupport.sheet(root, mutation.sheetId()), SnapshotMutationSupport.text(params, "sparklineId"));
            ranges.add(sourceRange(root, current));
            JsonNode patchRange = SnapshotMutationSupport.requiredObject(params, "patch").get("sourceRange");
            if (patchRange != null) ranges.add(SnapshotMutationSupport.range(root, patchRange));
        }
        return ranges.stream().distinct().toList();
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        switch (id()) {
            case "sparkline.add" -> add(root, sheet, mutation.sheetId(), params);
            case "sparkline.remove" -> remove(root, sheet, mutation.sheetId(), params);
            case "sparkline.update" -> update(root, sheet, mutation.sheetId(), params);
            case "sparkline.group.add", "sparkline.group.remove", "sparkline.group.replace" -> applyGroupState(sheet, mutation.sheetId(), params);
            default -> throw ServiceException.validation("Unsupported sparkline mutation: " + id());
        }
        return root;
    }

    private void add(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        ObjectNode input = SnapshotMutationSupport.requiredObject(params, "sparkline");
        ObjectNode sparkline = input.deepCopy();
        JsonNode groupState = params.get("groupState");
        if (groupState != null && !groupState.isNull()) {
            sparkline.remove("groupId");
            sparkline.remove("showAxis");
            sparkline.remove("showMarkers");
        }
        validateSparkline(root, sheetId, sparkline);
        ArrayNode sparklines = SnapshotMutationSupport.array(sheet, "sparklines");
        String id = SnapshotMutationSupport.text(sparkline, "id");
        if (SnapshotMutationSupport.findById(sparklines, id) != null) throw ServiceException.conflict("Sparkline already exists: " + id);
        sparklines.add(sparkline);
        if (groupState != null && !groupState.isNull()) {
            if (!groupState.isObject()) throw ServiceException.validation("Sparkline groupState must be an object");
            applyGroupState(sheet, sheetId, (ObjectNode) groupState);
        } else if (sparkline.has("groupId")) {
            throw ServiceException.validation("Sparkline group membership requires groupState");
        }
    }

    private void remove(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        String id = SnapshotMutationSupport.text(params, "sparklineId");
        ObjectNode current = sparkline(sheet, id);
        JsonNode groupState = params.get("groupState");
        if (current.has("groupId") && (groupState == null || groupState.isNull())) throw ServiceException.validation("Removing grouped sparkline requires groupState");
        if (groupState != null && !groupState.isNull()) {
            if (!groupState.isObject()) throw ServiceException.validation("Sparkline groupState must be an object");
            applyGroupState(sheet, sheetId, (ObjectNode) groupState);
        }
        SnapshotMutationSupport.removeById(SnapshotMutationSupport.array(sheet, "sparklines"), id);
    }

    private void update(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        ObjectNode current = sparkline(sheet, SnapshotMutationSupport.text(params, "sparklineId"));
        ObjectNode patch = SnapshotMutationSupport.requiredObject(params, "patch");
        ObjectNode next = current.deepCopy();
        patch.fields().forEachRemaining(entry -> next.set(entry.getKey(), entry.getValue().deepCopy()));
        next.put("sheetId", sheetId);
        validateSparkline(root, sheetId, next);
        SnapshotMutationSupport.array(sheet, "sparklines").set(SnapshotMutationSupport.indexById(SnapshotMutationSupport.array(sheet, "sparklines"), SnapshotMutationSupport.text(params, "sparklineId")), next);
    }

    private void applyGroupState(ObjectNode sheet, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.requireEntitySheet(params, sheetId);
        ArrayNode groupIds = SnapshotMutationSupport.requiredArray(params, "groupIds");
        ArrayNode groupEntries = SnapshotMutationSupport.requiredArray(params, "groups");
        ArrayNode members = SnapshotMutationSupport.requiredArray(params, "members");
        if (groupIds.size() > 10_000 || groupEntries.size() > 10_000 || members.size() > 10_000) throw ServiceException.validation("Sparkline group state is too large");
        Set<String> ids = new java.util.HashSet<>();
        for (JsonNode id : groupIds) {
            if (!id.isTextual() || id.asText().isBlank() || !ids.add(id.asText())) throw ServiceException.validation("Sparkline groupIds are invalid");
        }
        List<GroupEntry> nextGroups = new ArrayList<>();
        for (JsonNode raw : groupEntries) {
            if (!raw.isObject()) throw ServiceException.validation("Sparkline group entry must be an object");
            ObjectNode entry = (ObjectNode) raw;
            ObjectNode group = SnapshotMutationSupport.requiredObject(entry, "group");
            int index = integer(entry.get("index"), "Sparkline group index");
            validateGroup(sheetId, group);
            if (!ids.contains(SnapshotMutationSupport.text(group, "id"))) throw ServiceException.validation("Sparkline group is missing from groupIds");
            nextGroups.add(new GroupEntry(index, group));
        }
        ArrayNode groups = SnapshotMutationSupport.array(sheet, "sparklineGroups");
        ArrayNode retained = JsonNodeFactory.instance.arrayNode();
        for (JsonNode existing : groups) if (!ids.contains(existing.path("id").asText())) retained.add(existing);
        nextGroups.sort(Comparator.comparingInt(GroupEntry::index));
        for (GroupEntry entry : nextGroups) {
            int index = Math.max(0, Math.min(entry.index(), retained.size()));
            retained.insert(index, entry.group().deepCopy());
        }
        groups.removeAll();
        groups.addAll(retained);

        for (JsonNode raw : members) {
            if (!raw.isObject()) throw ServiceException.validation("Sparkline member state must be an object");
            ObjectNode member = (ObjectNode) raw;
            ObjectNode sparkline = sparkline(sheet, SnapshotMutationSupport.text(member, "sparklineId"));
            String type = SnapshotMutationSupport.text(member, "type");
            if (!Set.of("line", "column", "win-loss").contains(type)) throw ServiceException.validation("Sparkline type is invalid");
            sparkline.put("type", type);
            copyOptional(member, sparkline, "groupId");
            copyOptional(member, sparkline, "showAxis");
            copyOptional(member, sparkline, "showMarkers");
        }
        validateGroupMembership(sheet, sheetId);
    }

    private void validateSparkline(ObjectNode root, String sheetId, ObjectNode sparkline) {
        SnapshotMutationSupport.text(sparkline, "id");
        SnapshotMutationSupport.requireEntitySheet(sparkline, sheetId);
        String type = SnapshotMutationSupport.text(sparkline, "type");
        if (!Set.of("line", "column", "win-loss").contains(type)) throw ServiceException.validation("Sparkline type is invalid");
        ObjectNode anchor = SnapshotMutationSupport.requiredObject(sparkline, "anchor");
        SnapshotMutationSupport.coordinate(root, sheetId, anchor);
        sourceRange(root, sparkline);
        String groupId = SnapshotMutationSupport.optionalText(sparkline, "groupId");
        if (groupId != null) {
            ObjectNode group = SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(SnapshotMutationSupport.sheet(root, sheetId), "sparklineGroups"), groupId, "Sparkline group");
            boolean member = false;
            for (JsonNode id : SnapshotMutationSupport.requiredArray(group, "sparklineIds")) if (SnapshotMutationSupport.text(sparkline, "id").equals(id.asText())) member = true;
            if (!member) throw ServiceException.validation("Sparkline group membership is inconsistent");
        }
    }

    private RangeRef sourceRange(ObjectNode root, ObjectNode sparkline) {
        return SnapshotMutationSupport.range(root, sparkline.get("sourceRange"));
    }

    private ObjectNode sparkline(ObjectNode sheet, String id) {
        return SnapshotMutationSupport.requireById(SnapshotMutationSupport.array(sheet, "sparklines"), id, "Sparkline");
    }

    private void validateGroup(String sheetId, ObjectNode group) {
        SnapshotMutationSupport.text(group, "id");
        SnapshotMutationSupport.requireEntitySheet(group, sheetId);
        String type = SnapshotMutationSupport.text(group, "type");
        if (!Set.of("line", "column", "win-loss").contains(type)) throw ServiceException.validation("Sparkline group type is invalid");
        ArrayNode members = SnapshotMutationSupport.requiredArray(group, "sparklineIds");
        if (members.isEmpty()) throw ServiceException.validation("Sparkline group must have members");
        Set<String> memberIds = new java.util.HashSet<>();
        for (JsonNode member : members) if (!member.isTextual() || member.asText().isBlank() || !memberIds.add(member.asText())) throw ServiceException.validation("Sparkline group members are invalid");
    }

    private void validateGroupMembership(ObjectNode sheet, String sheetId) {
        ArrayNode groups = SnapshotMutationSupport.array(sheet, "sparklineGroups");
        for (JsonNode raw : groups) {
            if (!raw.isObject()) throw ServiceException.validation("Sparkline group is invalid");
            ObjectNode group = (ObjectNode) raw;
            validateGroup(sheetId, group);
            for (JsonNode member : group.path("sparklineIds")) {
                ObjectNode sparkline = sparkline(sheet, member.asText());
                if (!group.path("id").asText().equals(sparkline.path("groupId").asText())) throw ServiceException.validation("Sparkline group membership is inconsistent");
            }
        }
    }

    private void copyOptional(ObjectNode source, ObjectNode target, String key) {
        JsonNode value = source.get(key);
        if (value == null || value.isNull()) target.remove(key);
        else target.set(key, value.deepCopy());
    }

    private int integer(JsonNode value, String label) {
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 0) throw ServiceException.validation(label + " is invalid");
        return value.intValue();
    }

    private record GroupEntry(int index, ObjectNode group) {
    }
}
