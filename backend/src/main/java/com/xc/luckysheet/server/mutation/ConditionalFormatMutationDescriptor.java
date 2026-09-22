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

/** Canonical worksheet rule-order reducer shared by Home and contextual surfaces. */
final class ConditionalFormatMutationDescriptor extends CanonicalJsonMutationDescriptor {
    ConditionalFormatMutationDescriptor(String mutationId) {
        super(mutationId, WorkbookAclRole.EDITOR);
        if (!"cf.reorder".equals(mutationId)) throw new IllegalArgumentException("Unsupported conditional-format mutation: " + mutationId);
    }

    @Override
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        return List.of(SnapshotMutationSupport.wholeSheetRange(root, mutation.sheetId()));
    }

    @Override
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode sheet = SnapshotMutationSupport.sheet(root, mutation.sheetId());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        ArrayNode ruleIds = SnapshotMutationSupport.requiredArray(params, "ruleIds");
        ArrayNode rules = SnapshotMutationSupport.array(sheet, "conditionalFormats");
        Set<String> selected = new HashSet<>();
        ArrayNode reordered = rules.arrayNode();
        for (JsonNode value : ruleIds) {
            if (!value.isTextual() || !selected.add(value.textValue())) throw ServiceException.validation("cf.reorder ruleIds must be unique strings");
            ObjectNode rule = findRule(rules, value.textValue());
            if (rule == null) throw ServiceException.notFound("Conditional format rule not found: " + value.textValue());
            reordered.add(rule.deepCopy());
        }
        for (JsonNode value : rules) {
            if (value.isObject() && selected.add(value.path("id").asText())) reordered.add(value.deepCopy());
        }
        for (int index = 0; index < reordered.size(); index++) ((ObjectNode) reordered.get(index)).put("priority", index + 1);
        rules.removeAll();
        rules.addAll(reordered);
        return root;
    }

    private static ObjectNode findRule(ArrayNode rules, String id) {
        for (JsonNode value : rules) if (value.isObject() && id.equals(value.path("id").asText())) return (ObjectNode) value;
        return null;
    }
}
