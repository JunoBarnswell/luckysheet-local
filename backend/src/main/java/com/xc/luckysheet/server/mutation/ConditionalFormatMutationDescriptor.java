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
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/** Canonical reducers for conditional-format rule edits and ordering. */
final class ConditionalFormatMutationDescriptor extends CanonicalJsonMutationDescriptor {
    private static final Set<String> IDS = Set.of("cf.reorder", "cf.update");

    ConditionalFormatMutationDescriptor(String mutationId) {
        super(mutationId, WorkbookAclRole.EDITOR);
        if (!IDS.contains(mutationId)) throw new IllegalArgumentException("Unsupported conditional-format mutation: " + mutationId);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        return "cf.reorder".equals(id())
                ? reorderRanges(root, mutation.sheetId(), params)
                : updateRanges(root, mutation.sheetId(), params);
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        affectedRanges(snapshot, mutation);
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        if ("cf.reorder".equals(id())) reorder(sheet, params);
        else update(root, sheet, mutation.sheetId(), params);
        return root;
    }

    private List<RangeRef> reorderRanges(ObjectNode root, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.validateKnownKeys(params, Set.of("sheetId", "ruleIds", "ranges"), "cf.reorder");
        requireOwnerSheet(sheetId, params);
        ArrayNode rules = readRules(SnapshotMutationSupport.sheet(root, sheetId));
        Set<String> currentIds = new HashSet<>();
        List<RangeRef> expectedRanges = new ArrayList<>();
        for (JsonNode value : rules) {
            if (!value.isObject()) throw ServiceException.validation("conditionalFormats contains an invalid rule");
            ObjectNode rule = (ObjectNode) value;
            String ruleId = SnapshotMutationSupport.text(rule, "id");
            if (!currentIds.add(ruleId)) throw ServiceException.conflict("conditionalFormats contains duplicate rule identity: " + ruleId);
            expectedRanges.addAll(SnapshotMutationSupport.ranges(root, rule.get("ranges"), sheetId));
        }
        ArrayNode requestedIds = SnapshotMutationSupport.requiredArray(params, "ruleIds");
        Set<String> requested = new HashSet<>();
        for (JsonNode value : requestedIds) {
            if (!value.isTextual() || value.asText().isBlank() || !requested.add(value.textValue())) {
                throw ServiceException.validation("cf.reorder ruleIds must be unique, non-empty strings");
            }
        }
        if (requested.size() != currentIds.size() || !requested.equals(currentIds)) {
            throw ServiceException.validation("cf.reorder ruleIds must be a complete permutation of current rules");
        }
        return requireExactRanges(root, sheetId, params, expectedRanges, "cf.reorder");
    }

    private List<RangeRef> updateRanges(ObjectNode root, String sheetId, ObjectNode params) {
        SnapshotMutationSupport.validateKnownKeys(params, Set.of("sheetId", "before", "after", "ranges"), "cf.update");
        requireOwnerSheet(sheetId, params);
        ObjectNode before = SnapshotMutationSupport.requiredObject(params, "before");
        ObjectNode after = SnapshotMutationSupport.requiredObject(params, "after");
        SheetRuleLifecycle.validateRule(root, sheetId, before, "conditionalFormats");
        SheetRuleLifecycle.validateRule(root, sheetId, after, "conditionalFormats");
        String ruleId = SnapshotMutationSupport.text(before, "id");
        if (!ruleId.equals(SnapshotMutationSupport.text(after, "id"))) {
            throw ServiceException.validation("cf.update cannot change rule identity");
        }
        ArrayNode rules = readRules(SnapshotMutationSupport.sheet(root, sheetId));
        ObjectNode current = findUniqueRule(rules, ruleId);
        if (current == null || !current.equals(before)) {
            throw ServiceException.conflict("Conditional format changed before update: " + ruleId);
        }
        List<RangeRef> expectedRanges = new ArrayList<>(SnapshotMutationSupport.ranges(root, before.get("ranges"), sheetId));
        expectedRanges.addAll(SnapshotMutationSupport.ranges(root, after.get("ranges"), sheetId));
        return requireExactRanges(root, sheetId, params, expectedRanges, "cf.update");
    }

    private List<RangeRef> requireExactRanges(
            ObjectNode root,
            String sheetId,
            ObjectNode params,
            List<RangeRef> expected,
            String mutationId
    ) {
        JsonNode rawRanges = params.get("ranges");
        if (expected.isEmpty()) {
            if (rawRanges == null || !rawRanges.isArray() || !rawRanges.isEmpty()) {
                throw ServiceException.validation(mutationId + " ranges must exactly match the affected rule ranges");
            }
            return List.of();
        }
        List<RangeRef> declared = SnapshotMutationSupport.ranges(root, params.get("ranges"), sheetId);
        if (!rangeMultisetsEqual(declared, expected)) {
            throw ServiceException.validation(mutationId + " ranges must exactly match the affected rule ranges");
        }
        return List.copyOf(expected);
    }

    private boolean rangeMultisetsEqual(List<RangeRef> left, List<RangeRef> right) {
        if (left.size() != right.size()) return false;
        Map<RangeRef, Integer> counts = new HashMap<>();
        for (RangeRef range : left) counts.merge(range, 1, Integer::sum);
        for (RangeRef range : right) {
            Integer count = counts.get(range);
            if (count == null) return false;
            if (count == 1) counts.remove(range);
            else counts.put(range, count - 1);
        }
        return counts.isEmpty();
    }

    private void requireOwnerSheet(String sheetId, ObjectNode params) {
        if (!sheetId.equals(SnapshotMutationSupport.text(params, "sheetId"))) {
            throw ServiceException.validation(id() + " owner sheet does not match its mutation envelope");
        }
    }

    private void reorder(ObjectNode sheet, ObjectNode params) {
        ArrayNode rules = SnapshotMutationSupport.array(sheet, "conditionalFormats");
        Map<String, ObjectNode> byId = new HashMap<>();
        for (JsonNode value : rules) {
            if (!value.isObject()) throw ServiceException.validation("conditionalFormats contains an invalid rule");
            ObjectNode rule = (ObjectNode) value;
            String ruleId = SnapshotMutationSupport.text(rule, "id");
            if (byId.putIfAbsent(ruleId, rule) != null) {
                throw ServiceException.conflict("conditionalFormats contains duplicate rule identity: " + ruleId);
            }
        }
        ArrayNode requestedIds = SnapshotMutationSupport.requiredArray(params, "ruleIds");
        ArrayNode reordered = rules.arrayNode();
        for (JsonNode value : requestedIds) {
            ObjectNode rule = byId.get(value.textValue());
            if (rule == null) throw ServiceException.notFound("Conditional format rule not found: " + value.textValue());
            reordered.add(rule.deepCopy());
        }
        for (int index = 0; index < reordered.size(); index++) ((ObjectNode) reordered.get(index)).put("priority", index + 1);
        rules.removeAll();
        rules.addAll(reordered);
    }

    private void update(ObjectNode root, ObjectNode sheet, String sheetId, ObjectNode params) {
        String ruleId = SnapshotMutationSupport.text(SnapshotMutationSupport.requiredObject(params, "before"), "id");
        ObjectNode before = SnapshotMutationSupport.requiredObject(params, "before");
        ObjectNode after = SnapshotMutationSupport.requiredObject(params, "after");
        ArrayNode rules = SnapshotMutationSupport.array(sheet, "conditionalFormats");
        ObjectNode current = findUniqueRule(rules, ruleId);
        if (current == null || !current.equals(before)) throw ServiceException.conflict("Conditional format changed before update: " + ruleId);
        SheetRuleLifecycle.validateRule(root, sheetId, after, "conditionalFormats");
        for (int index = 0; index < rules.size(); index++) {
            if (rules.get(index) == current) {
                rules.set(index, after.deepCopy());
                return;
            }
        }
        throw ServiceException.conflict("Conditional format disappeared during update: " + ruleId);
    }

    private static ArrayNode readRules(ObjectNode sheet) {
        JsonNode value = sheet.get("conditionalFormats");
        if (value == null || value.isNull()) return JsonNodeFactory.instance.arrayNode();
        if (!value.isArray()) throw ServiceException.validation("conditionalFormats must be an array");
        return (ArrayNode) value;
    }

    private static ObjectNode findUniqueRule(ArrayNode rules, String id) {
        ObjectNode found = null;
        for (JsonNode value : rules) {
            if (!value.isObject()) throw ServiceException.validation("conditionalFormats contains an invalid rule");
            ObjectNode rule = (ObjectNode) value;
            String currentId = SnapshotMutationSupport.text(rule, "id");
            if (!id.equals(currentId)) continue;
            if (found != null) throw ServiceException.conflict("conditionalFormats contains duplicate rule identity: " + id);
            found = rule;
        }
        return found;
    }
}
