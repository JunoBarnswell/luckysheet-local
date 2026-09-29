package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.StructuralStateChanges;
import com.xc.luckysheet.server.contract.StructuralStateChanges.CellChange;
import com.xc.luckysheet.server.contract.StructuralStateChanges.CellRowChange;
import com.xc.luckysheet.server.contract.StructuralStateChanges.PropertyChange;
import com.xc.luckysheet.server.contract.StructuralStateChanges.PropertyValue;
import com.xc.luckysheet.server.contract.StructuralStateChanges.SheetChange;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class StructuralStateChangeReplayTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void migratedJavaAxisFactsReplayAndInvertWithoutReenteringStructuralReducers() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        for (String id : List.of("rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted")) {
            for (int at : List.of(0, 1, 4)) {
                ObjectNode before = workbook();
                JsonNode original = before.deepCopy();
                ObjectNode params = mapper.createObjectNode().put("at", at).put("count", 2);
                OperationMutation intent = new OperationMutation(id, "data", params);
                JsonNode after = registry.require(id, false).applyWithPatch(before, intent).snapshot();
                StructuralStateChanges changes = StructuralStateChangeReplay.captureForMigration(before, after);
                StructuralStateChanges wire = mapper.readValue(mapper.writeValueAsBytes(changes), StructuralStateChanges.class);
                assertEquals(after, StructuralStateChangeReplay.apply(before, wire), id + " apply at " + at);
                assertEquals(original, StructuralStateChangeReplay.apply(after, wire.inverse()), id + " undo at " + at);
                JsonNode undone = StructuralStateChangeReplay.apply(after, wire.inverse());
                assertEquals(after, StructuralStateChangeReplay.apply(undone, wire), id + " redo at " + at);
                assertEquals(original, before, "Planning and replay must not mutate the input");
            }
        }
    }

    @Test
    void deletionFactsRetainFormulaCachesMetadataAndNonFormulaCellsExactly() throws Exception {
        ObjectNode before = workbook();
        JsonNode after = new MutationDescriptorRegistry().require("rows.deleted", false).apply(before,
                new OperationMutation("rows.deleted", "data", mapper.createObjectNode().put("at", 1).put("count", 1)));
        StructuralStateChanges changes = StructuralStateChangeReplay.captureForMigration(before, after);
        assertFalse(changes.sheets().isEmpty());
        assertEquals(before, StructuralStateChangeReplay.apply(after, changes.inverse()));
        assertEquals("='Data Sheet'!A2", before.path("sheets").get(1).path("cells").path("0").path("0").path("formula").asText());
    }

    @Test
    void lateCellOrMetadataMismatchRejectsBeforeAnyOwnedSnapshotWrite() throws Exception {
        ObjectNode before = workbook();
        ObjectNode after = before.deepCopy();
        after.put("name", "renamed");
        cell(after, 0, 0, 0).put("value", 99);
        cell(after, 1, 0, 0).put("formula", "=42");
        StructuralStateChanges changes = StructuralStateChangeReplay.captureForMigration(before, after);

        ObjectNode stale = before.deepCopy();
        cell(stale, 1, 0, 0).put("formula", "=43");
        JsonNode untouched = stale.deepCopy();
        ServiceException rejected = assertThrows(ServiceException.class,
                () -> StructuralStateChangeReplay.applyOnOwnedSnapshot(stale, changes));
        assertEquals("STRUCTURAL_PATCH_PRECONDITION", rejected.code());
        assertEquals(untouched, stale);

        ObjectNode staleName = before.deepCopy();
        staleName.put("name", "concurrent");
        JsonNode untouchedName = staleName.deepCopy();
        assertThrows(ServiceException.class, () -> StructuralStateChangeReplay.applyOnOwnedSnapshot(staleName, changes));
        assertEquals(untouchedName, staleName);
    }

    @Test
    void emptyRowsExplicitNullAndAbsentMetadataRemainDistinctAcrossWireAndInverse() throws Exception {
        ObjectNode before = workbook();
        ObjectNode after = before.deepCopy();
        sheet(before, 0).putNull("tabColor");
        ((ObjectNode) sheet(before, 0).get("cells")).set("9", mapper.createObjectNode());
        sheet(after, 0).remove("tabColor");
        sheet(after, 1).putNull("tabColor");
        ((ObjectNode) sheet(after, 1).get("cells")).set("8", mapper.createObjectNode());
        StructuralStateChanges changes = mapper.readValue(mapper.writeValueAsBytes(
                StructuralStateChangeReplay.captureForMigration(before, after)), StructuralStateChanges.class);
        assertEquals(after, StructuralStateChangeReplay.apply(before, changes));
        assertEquals(before, StructuralStateChangeReplay.apply(after, changes.inverse()));
    }

    @Test
    void sheetLifecycleAndReorderingPreserveEveryRecordedPropertyAndCell() throws Exception {
        ObjectNode before = workbook();
        ObjectNode after = before.deepCopy();
        ObjectNode created = sheet(before, 0).deepCopy();
        created.put("id", "new").put("name", "New");
        ObjectNode retained = sheet(after, 1);
        ArrayNode sheets = mapper.createArrayNode().add(created).add(retained);
        after.set("sheets", sheets);
        StructuralStateChanges changes = StructuralStateChangeReplay.captureForMigration(before, after);
        assertEquals(List.of("new", "summary"), changes.sheetOrderAfter());
        assertEquals(after, StructuralStateChangeReplay.apply(before, changes));
        assertEquals(before, StructuralStateChangeReplay.apply(after, changes.inverse()));
    }

    @Test
    void sheetAndRowDeletionRejectOmittedStoredFactsInsteadOfDiscardingData() throws Exception {
        ObjectNode before = workbook();
        ObjectNode after = before.deepCopy();
        ((ArrayNode) after.get("sheets")).remove(0);
        StructuralStateChanges complete = StructuralStateChangeReplay.captureForMigration(before, after);
        SheetChange deleted = complete.sheets().get(0);
        SheetChange missingRows = new SheetChange(deleted.sheetId(), true, false, deleted.properties(), List.of());
        StructuralStateChanges malformed = new StructuralStateChanges(complete.unitId(), complete.sheetOrderBefore(),
                complete.sheetOrderAfter(), complete.workbookProperties(), List.of(missingRows));
        JsonNode untouched = before.deepCopy();
        assertThrows(ServiceException.class, () -> StructuralStateChangeReplay.applyOnOwnedSnapshot(before, malformed));
        assertEquals(untouched, before);

        ObjectNode rowDeleted = before.deepCopy();
        ((ObjectNode) sheet(rowDeleted, 0).get("cells")).remove("1");
        StructuralStateChanges fullRow = StructuralStateChangeReplay.captureForMigration(before, rowDeleted);
        SheetChange target = fullRow.sheets().get(0);
        CellRowChange row = target.rows().get(0);
        CellRowChange partialRow = new CellRowChange(row.row(), true, false, row.cells().subList(0, 1));
        StructuralStateChanges incompleteRow = new StructuralStateChanges(fullRow.unitId(), fullRow.sheetOrderBefore(),
                fullRow.sheetOrderAfter(), List.of(), List.of(new SheetChange(target.sheetId(), true, true, List.of(), List.of(partialRow))));
        assertThrows(ServiceException.class, () -> StructuralStateChangeReplay.applyOnOwnedSnapshot(before, incompleteRow));
        assertEquals(untouched, before);
    }

    @Test
    void ownedReplayDoesNotCloneOrRewriteUnchangedWorksheetOrCellPayloads() throws Exception {
        ObjectNode before = workbook();
        ObjectNode after = before.deepCopy();
        cell(after, 0, 0, 0).put("value", 101);
        StructuralStateChanges changes = StructuralStateChangeReplay.captureForMigration(before, after);
        ObjectNode owned = before.deepCopy();
        JsonNode untouchedSheet = owned.path("sheets").get(1);
        JsonNode untouchedCell = owned.path("sheets").get(0).path("cells").path("1").path("1");
        StructuralStateChangeReplay.applyOnOwnedSnapshot(owned, changes);
        assertSame(untouchedSheet, owned.path("sheets").get(1));
        assertSame(untouchedCell, owned.path("sheets").get(0).path("cells").path("1").path("1"));
        assertEquals(after, owned);
        assertNotSame(changes.sheets().get(0).rows().get(0).cells().get(0).after(), cell(owned, 0, 0, 0));
    }

    @Test
    void fullWorksheetDimensionsWithFewStoredCellsProduceOnlySparseChangedFacts() throws Exception {
        ObjectNode before = workbook();
        sheet(before, 0).put("rowCount", 1_048_576).put("columnCount", 16_384);
        ObjectNode after = before.deepCopy();
        cell(after, 0, 1_048_575, 16_383).put("value", "edge");
        StructuralStateChanges changes = StructuralStateChangeReplay.captureForMigration(before, after);
        assertEquals(1, changes.sheets().size());
        assertEquals(1, changes.sheets().get(0).rows().size());
        assertEquals(1, changes.sheets().get(0).rows().get(0).cells().size());
        assertTrue(mapper.writeValueAsBytes(changes).length < 1_000);
        assertEquals(after, StructuralStateChangeReplay.apply(before, changes));
        assertEquals(before, StructuralStateChangeReplay.apply(after, changes.inverse()));
    }

    @Test
    void contractRejectsDuplicateOwnersIdentityRewritesAndUnrecordedLifecycle() throws Exception {
        ObjectNode value = mapper.createObjectNode().put("value", 1);
        CellChange cell = new CellChange(0, null, value);
        assertThrows(IllegalArgumentException.class, () -> new CellRowChange(0, false, true, List.of(cell, cell)));
        assertThrows(IllegalArgumentException.class, () -> new CellChange(16_384, null, value));
        PropertyChange identity = new PropertyChange("id", PropertyValue.of(mapper.getNodeFactory().textNode("data")),
                PropertyValue.of(mapper.getNodeFactory().textNode("renamed")));
        assertThrows(IllegalArgumentException.class, () -> new SheetChange("data", true, true, List.of(identity), List.of()));
        assertThrows(IllegalArgumentException.class, () -> new StructuralStateChanges("book", List.of("data"),
                List.of("summary"), List.of(), List.of()));
        ObjectNode before = workbook();
        ObjectNode after = before.deepCopy();
        after.put("unitId", "other");
        assertThrows(ServiceException.class, () -> StructuralStateChangeReplay.captureForMigration(before, after));
    }

    @Test
    void wireRejectsMissingOrNullPresenceAndCoordinatesInsteadOfAssumingFalseOrZero() {
        assertThrows(com.fasterxml.jackson.core.JsonProcessingException.class,
                () -> mapper.readValue("{\"value\":null}", PropertyValue.class));
        assertThrows(com.fasterxml.jackson.core.JsonProcessingException.class,
                () -> mapper.readValue("{\"present\":null,\"value\":null}", PropertyValue.class));
        assertThrows(com.fasterxml.jackson.core.JsonProcessingException.class,
                () -> mapper.readValue("{\"before\":null,\"after\":{\"value\":1}}", CellChange.class));
        assertThrows(com.fasterxml.jackson.core.JsonProcessingException.class,
                () -> mapper.readValue("{\"column\":null,\"before\":null,\"after\":{\"value\":1}}", CellChange.class));
        assertThrows(com.fasterxml.jackson.core.JsonProcessingException.class,
                () -> mapper.readValue("{\"row\":null,\"beforePresent\":false,\"afterPresent\":true,\"cells\":[]}", CellRowChange.class));
    }

    private ObjectNode workbook() throws Exception {
        return (ObjectNode) mapper.readTree("""
                {"schema":"WorkbookSnapshot","version":10,"unitId":"book","name":"Facts",
                 "definedNameModels":[],"definedNames":{},
                 "sheets":[
                   {"id":"data","name":"Data Sheet","rowCount":20,"columnCount":10,
                    "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                    "rowHeightsPx":{"1":28},"columnWidthsPx":{},"hiddenRows":[2],"hiddenColumns":[3],
                    "merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],
                    "drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],
                    "protectionRules":[],"outline":{"groups":[]},
                    "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                    "cells":{"0":{"0":{"value":11}},
                      "1":{"0":{"value":12,"style":{"fontFamily":"Arial"}},"1":{"value":"kept"}},
                      "5":{"4":{"value":15}}}},
                   {"id":"summary","name":"Summary","rowCount":20,"columnCount":10,
                    "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                    "rowHeightsPx":{},"columnWidthsPx":{},"hiddenRows":[],"hiddenColumns":[],
                    "merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],
                    "drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],
                    "protectionRules":[],"outline":{"groups":[]},
                    "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                    "cells":{"0":{"0":{"value":12,"formula":"='Data Sheet'!A2"}}}}
                 ]}
                """);
    }

    private static ObjectNode sheet(ObjectNode workbook, int index) {
        return (ObjectNode) workbook.path("sheets").get(index);
    }

    private static ObjectNode cell(ObjectNode workbook, int sheet, int row, int column) {
        return SnapshotMutationSupport.cell(sheet(workbook, sheet), new SnapshotMutationSupport.CellCoordinate(row, column), true);
    }
}
