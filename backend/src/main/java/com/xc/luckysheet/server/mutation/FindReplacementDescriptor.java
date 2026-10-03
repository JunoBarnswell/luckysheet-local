package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.service.ServiceException;
import java.util.*;

/** Exact preimage-checked replacement shared by forward commit and owned undo. */
public final class FindReplacementDescriptor extends CanonicalJsonMutationDescriptor {
    public FindReplacementDescriptor() { super("find.replaced", WorkbookRole.EDITOR); }
    @Override public List<RangeRef> affectedRanges(JsonNode snapshot, OperationMutation mutation) {
        ObjectNode root = SnapshotMutationSupport.root(snapshot);
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        if (!Set.of("forward", "reverse").contains(params.path("direction").asText())) throw ServiceException.validation("Invalid replacement direction");
        JsonNode patches = params.path("patches");
        if (!patches.isArray() || patches.isEmpty() || patches.size() > SnapshotMutationSupport.MAX_CHANGED_CELLS) throw ServiceException.validation("Replacement patches must be bounded");
        List<RangeRef> ranges = new ArrayList<>(); Set<PatchIdentity> identities = new HashSet<>();
        for (JsonNode patch : patches) {
            String kind = patch.path("kind").asText();
            if (!Set.of("cell", "note", "comment").contains(kind)) throw ServiceException.validation("Invalid replacement patch kind");
            JsonNode match = patch.path("match"); String sheetId = match.path("sheetId").asText();
            ObjectNode address = matchObject(match);
            var coordinate = SnapshotMutationSupport.coordinate(root, sheetId, address);
            if (!identities.add(new PatchIdentity(sheetId, coordinate.row(), coordinate.column(), kind, match.path("sourceId").asText()))) throw ServiceException.validation("Duplicate replacement patch");
            ranges.add(new RangeRef(sheetId, coordinate.row(), coordinate.row(), coordinate.column(), coordinate.column()));
        }
        return ranges;
    }
    @Override public JsonNode apply(JsonNode snapshot, OperationMutation mutation) {
        affectedRanges(snapshot, mutation);
        ObjectNode root = SnapshotMutationSupport.root(snapshot.deepCopy());
        ObjectNode params = SnapshotMutationSupport.params(mutation);
        boolean forward = "forward".equals(params.path("direction").asText());
        for (JsonNode patch : params.path("patches")) {
            ObjectNode match = matchObject(patch.path("match"));
            ObjectNode sheet = SnapshotMutationSupport.sheet(root, match.path("sheetId").asText());
            var coordinate = SnapshotMutationSupport.coordinate(root, match.path("sheetId").asText(), match);
            String kind = patch.path("kind").asText();
            JsonNode expected = patch.get(forward ? "previous" : "next");
            JsonNode after = patch.get(forward ? "next" : "previous");
            if ("cell".equals(kind)) {
                JsonNode current = SnapshotMutationSupport.cell(sheet, coordinate, false);
                assertPreimage(current, expected);
                if (after == null || after.isNull()) SnapshotMutationSupport.removeCell(sheet, coordinate);
                else { if (!after.isObject() || !after.has("value")) throw ServiceException.validation("Replacement cell is invalid"); SnapshotMutationSupport.putCell(sheet, coordinate, after); }
            } else if ("note".equals(kind)) {
                assertPreimage(SnapshotMutationSupport.findNote(sheet, coordinate), expected);
                if (after == null || after.isNull()) SnapshotMutationSupport.removeNote(sheet, coordinate);
                else { if (!after.path("text").isTextual() || after.path("id").asText().isBlank()) throw ServiceException.validation("Replacement note is invalid"); SnapshotMutationSupport.putNote(sheet, coordinate, after); }
            } else {
                JsonNode node = sheet.path("review").path("threadsById").get(match.path("sourceId").asText());
                if (!(node instanceof ObjectNode thread) || thread.path("row").asInt(-1) != coordinate.row() || thread.path("column").asInt(-1) != coordinate.column()) throw ServiceException.conflict("Replacement comment changed");
                assertPreimage(thread.get("text"), patch.get(forward ? "previousText" : "nextText"));
                JsonNode text = patch.get(forward ? "nextText" : "previousText");
                if (text == null || !text.isTextual()) throw ServiceException.validation("Replacement comment text is invalid");
                thread.set("text", text.deepCopy());
            }
        }
        return root;
    }
    private record PatchIdentity(String sheet, int row, int column, String kind, String source) {}
    private static ObjectNode matchObject(JsonNode match) {
        if (!(match instanceof ObjectNode object)) throw ServiceException.validation("Replacement match must be an object");
        return object;
    }
    private static void assertPreimage(JsonNode current, JsonNode expected) {
        if (expected == null || expected.isNull()) { if (current != null && !current.isNull()) throw ServiceException.conflict("Replacement source changed"); }
        else if (!expected.equals(current)) throw ServiceException.conflict("Replacement source changed");
    }
}
