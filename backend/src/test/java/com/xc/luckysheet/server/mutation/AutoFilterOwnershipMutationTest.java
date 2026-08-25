package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class AutoFilterOwnershipMutationTest {
    private final ObjectMapper mapper = new ObjectMapper();
    private final MutationDescriptorRegistry registry = new MutationDescriptorRegistry();

    @Test
    void worksheetFilterRejectsSameContainedAndPartialTableOverlapsBeforePrepare() {
        for (int[] candidate : List.of(
                new int[]{0, 4, 0, 2},
                new int[]{1, 3, 1, 1},
                new int[]{2, 6, 2, 4})) {
            ObjectNode snapshot = snapshotWithTableFilter(0, 4, 0, 2);
            OperationMutation mutation = worksheetFilter(candidate[0], candidate[1], candidate[2], candidate[3]);
            assertRejectedWithoutMutation(snapshot, mutation);
        }
    }

    @Test
    void worksheetFilterAcceptsDisjointTableRangeAndReplayProducesTheSameOwnerState() {
        ObjectNode snapshot = snapshotWithTableFilter(0, 4, 0, 2);
        OperationMutation mutation = worksheetFilter(6, 9, 0, 2);
        assertDoesNotThrow(() -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));
        JsonNode next = registry.applyPublicMutations(snapshot, List.of(mutation));
        assertEquals(6, next.path("sheets").get(0).path("autoFilter").path("range").path("startRow").asInt());
        assertEquals(0, next.path("sheets").get(0).path("sheetTables").get(0).path("autoFilter").path("range").path("startRow").asInt());
    }

    @Test
    void tableFilterRejectsWorksheetAndOtherTableOverlapsButAllowsDisjointOwner() {
        ObjectNode snapshot = snapshotWithWorksheetFilter(0, 4, 0, 2);
        ObjectNode table = table("table-1", 2, 6, 0, 2, false);
        ((ObjectNode) snapshot.path("sheets").get(0)).withArray("sheetTables").add(table);

        assertRejectedWithoutMutation(snapshot, tableFilter("table-1", 2, 6, 0, 2));

        ObjectNode tableOverlap = baseSnapshot();
        var overlapSheet = (ObjectNode) tableOverlap.path("sheets").get(0);
        overlapSheet.withArray("sheetTables").add(table("table-1", 2, 6, 0, 2, false));
        overlapSheet.withArray("sheetTables").add(table("table-2", 4, 8, 0, 2, true));
        assertRejectedWithoutMutation(tableOverlap, tableFilter("table-1", 2, 6, 0, 2));

        ObjectNode disjointSnapshot = snapshotWithWorksheetFilter(0, 4, 0, 2);
        ((ObjectNode) disjointSnapshot.path("sheets").get(0)).withArray("sheetTables").add(table("table-1", 6, 9, 0, 2, false));
        OperationMutation disjoint = tableFilter("table-1", 6, 9, 0, 2);
        assertDoesNotThrow(() -> registry.prepare(disjointSnapshot, disjoint, WorkbookAclRole.EDITOR));
        JsonNode next = registry.applyPublicMutations(disjointSnapshot, List.of(disjoint));
        assertEquals(6, next.path("sheets").get(0).path("sheetTables").get(0).path("autoFilter").path("range").path("startRow").asInt());

    }

    @Test
    void tableAddAndResizeRejectEmbeddedOrMovedFilterOverlap() {
        ObjectNode snapshot = snapshotWithWorksheetFilter(0, 4, 0, 2);
        OperationMutation add = new OperationMutation("sheetTable.add", "sheet-1", table("table-1", 2, 6, 0, 2, true));
        assertRejectedWithoutMutation(snapshot, add);

        ObjectNode resizeSnapshot = snapshotWithWorksheetFilter(0, 4, 0, 2);
        ((ObjectNode) resizeSnapshot.path("sheets").get(0)).withArray("sheetTables").add(table("table-1", 6, 9, 0, 2, true));
        ObjectNode updated = table("table-1", 2, 6, 0, 2, true);
        OperationMutation resize = new OperationMutation("sheetTable.update", "sheet-1", updated);
        assertRejectedWithoutMutation(resizeSnapshot, resize);
    }

    @Test
    void sameOwnerReplacementAndExplicitRemovalRemainAtomic() {
        ObjectNode snapshot = snapshotWithTableFilter(0, 4, 0, 2);
        OperationMutation replacement = tableFilter("table-1", 0, 4, 0, 2);
        assertDoesNotThrow(() -> registry.prepare(snapshot, replacement, WorkbookAclRole.EDITOR));
        JsonNode replaced = registry.applyPublicMutations(snapshot, List.of(replacement));
        assertEquals(0, replaced.path("sheets").get(0).path("sheetTables").get(0).path("autoFilter").path("range").path("startRow").asInt());

        ObjectNode params = mapper.createObjectNode().put("sheetId", "sheet-1").put("tableId", "table-1");
        OperationMutation remove = new OperationMutation("sheetTable.autoFilter.set", "sheet-1", params);
        JsonNode removed = registry.applyPublicMutations(replaced, List.of(remove));
        assertEquals(0, removed.path("sheets").get(0).path("sheetTables").get(0).path("autoFilter").size());
    }

    private void assertRejectedWithoutMutation(ObjectNode snapshot, OperationMutation mutation) {
        JsonNode before = snapshot.deepCopy();
        ServiceException prepared = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", prepared.code());
        assertEquals(before, snapshot);
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(mutation)));
        assertEquals(before, snapshot);
    }

    private ObjectNode snapshotWithWorksheetFilter(int startRow, int endRow, int startColumn, int endColumn) {
        ObjectNode snapshot = baseSnapshot();
        ((ObjectNode) snapshot.path("sheets").get(0)).set("autoFilter", filter(startRow, endRow, startColumn, endColumn));
        return snapshot;
    }

    private ObjectNode snapshotWithTableFilter(int startRow, int endRow, int startColumn, int endColumn) {
        ObjectNode snapshot = baseSnapshot();
        ((ObjectNode) snapshot.path("sheets").get(0)).withArray("sheetTables").add(table("table-1", startRow, endRow, startColumn, endColumn, true));
        return snapshot;
    }

    private ObjectNode baseSnapshot() {
        ObjectNode snapshot = mapper.createObjectNode();
        ObjectNode sheet = snapshot.putArray("sheets").addObject();
        sheet.put("id", "sheet-1").put("rowCount", 20).put("columnCount", 10);
        sheet.putObject("cells");
        sheet.putArray("sheetTables");
        return snapshot;
    }

    private OperationMutation worksheetFilter(int startRow, int endRow, int startColumn, int endColumn) {
        ObjectNode params = mapper.createObjectNode().put("sheetId", "sheet-1");
        params.set("autoFilter", filter(startRow, endRow, startColumn, endColumn));
        return new OperationMutation("autoFilter.set", "sheet-1", params);
    }

    private OperationMutation tableFilter(String tableId, int startRow, int endRow, int startColumn, int endColumn) {
        ObjectNode params = mapper.createObjectNode().put("sheetId", "sheet-1").put("tableId", tableId);
        params.set("autoFilter", filter(startRow, endRow, startColumn, endColumn));
        return new OperationMutation("sheetTable.autoFilter.set", "sheet-1", params);
    }

    private ObjectNode filter(int startRow, int endRow, int startColumn, int endColumn) {
        ObjectNode filter = mapper.createObjectNode().put("sheetId", "sheet-1");
        filter.set("range", range(startRow, endRow, startColumn, endColumn));
        filter.putObject("columns");
        return filter;
    }

    private ObjectNode table(String id, int startRow, int endRow, int startColumn, int endColumn, boolean withFilter) {
        ObjectNode table = mapper.createObjectNode()
                .put("id", id).put("sheetId", "sheet-1").put("name", id)
                .put("hasHeaderRow", true).put("hasTotalRow", false)
                .put("showBandedRows", true).put("showBandedColumns", false)
                .put("showFirstColumn", false).put("showLastColumn", false)
                .put("showFilterButton", true).put("autoExpand", "both");
        table.set("range", range(startRow, endRow, startColumn, endColumn));
        int width = endColumn - startColumn + 1;
        var columns = table.putArray("columns");
        for (int index = 0; index < width; index++) columns.addObject().put("id", id + "-column-" + index).put("name", "Column" + index);
        if (withFilter) table.set("autoFilter", filter(startRow, endRow, startColumn, endColumn));
        return table;
    }

    private ObjectNode range(int startRow, int endRow, int startColumn, int endColumn) {
        return mapper.createObjectNode().put("sheetId", "sheet-1")
                .put("startRow", startRow).put("endRow", endRow)
                .put("startColumn", startColumn).put("endColumn", endColumn);
    }
}
