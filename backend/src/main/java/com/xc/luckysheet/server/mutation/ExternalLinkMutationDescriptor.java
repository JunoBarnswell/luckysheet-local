package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.*;
import com.xc.luckysheet.server.service.ServiceException;
import java.util.List;
import java.util.Set;

final class ExternalLinkMutationDescriptor extends CanonicalJsonMutationDescriptor {
    static final Set<String> IDS = Set.of("externalLink.set", "externalLink.remove");
    ExternalLinkMutationDescriptor(String id) { super(id, WorkbookRole.EDITOR); }
    public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) { validate(mutation); return List.of(); }
    private ObjectNode validate(OperationMutation mutation) {
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        SnapshotMutationSupport.validateKnownKeys(params, id().equals("externalLink.set") ? Set.of("link") : Set.of("linkId"), id());
        if (id().equals("externalLink.set")) ExternalLinkDefinitionValidator.validate(params.get("link"));
        else SnapshotMutationSupport.text(params, "linkId");
        return params;
    }
    public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = validate(mutation);
        var links = SnapshotMutationSupport.array(SnapshotMutationSupport.object(root, "dataModel"), "externalLinks");
        String id = id().equals("externalLink.set") ? params.path("link").path("id").asText() : params.path("linkId").asText();
        int index = SnapshotMutationSupport.indexById(links, id);
        if (id().equals("externalLink.remove")) {
            if (index < 0) throw ServiceException.notFound("External link not found");
            links.remove(index);
        } else {
            if (params.path("link").path("sourceUnitId").asText().equals(root.path("unitId").asText())) throw ServiceException.validation("External link source must be another workbook");
            if (index < 0) links.add(params.get("link").deepCopy()); else links.set(index, params.get("link").deepCopy());
        }
        ExternalLinkDefinitionValidator.validateCollection(links);
        return root;
    }
}
