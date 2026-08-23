package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class MutationDescriptorRegistryTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void unknownMutationsFailClosed() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ServiceException error = assertThrows(ServiceException.class, () -> registry.require("unknown.mutation", false));
        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void cellSetUsesServerResolvedRangeAndChangesSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("{\"sheets\":[{\"id\":\"sheet-1\",\"cells\":{}}]}");
        var mutation = new OperationMutation("cell.set", "sheet-1", mapper.readTree("{\"row\":2,\"column\":3,\"value\":{\"value\":42}}"));
        assertEquals(2, registry.resolveRanges(snapshot, mutation).get(0).startRow());
        var next = registry.applyPublicMutations(snapshot, List.of(mutation));
        assertEquals(42, next.path("sheets").get(0).path("cells").path("2").path("3").path("value").asInt());
    }

    @Test
    void internalRestoreCannotBeSubmittedByClient() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        assertThrows(ServiceException.class, () -> registry.require("workbook.restore", false));
    }

    @Test
    void knownMutationWithoutServerReducerFailsClosedInsteadOfBeingStoredAsOpaqueJson() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ServiceException error = assertThrows(ServiceException.class, () -> registry.require("sheet.duplicated", false));
        assertEquals("SERVICE_UNAVAILABLE", error.code());
        assertEquals(true, registry.ids().contains("sheet.duplicated"));
        assertEquals(false, registry.acceptedIds().contains("sheet.duplicated"));
    }

    @Test
    void acceptedMutationSurfaceIsExplicitAndAllOtherKnownMutationsRemainFailClosed() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        assertEquals(Set.of(
                "cell.set", "cell.restore", "range.set", "range.paste", "range.clear", "range.clear.restore",
                "style.set", "merge.set", "merge.remove", "freeze.set", "row.resize", "column.resize", "view.set", "sheet.hidden", "sheet.unhidden", "sheet.tabColor",
                "note.set", "note.remove", "note.visibility", "comment.add", "comment.reply", "comment.reply.remove", "comment.resolve", "comment.remove",
                "sheet.protect.set", "sheet.protect.remove", "workbook.renamed",
                "sheet.reordered",
                "row.hidden", "row.unhidden", "rows.unhidden.all", "rows.hidden.restore",
                "column.hidden", "column.unhidden", "columns.unhidden.all", "columns.hidden.restore",
                "filter.set", "filter.remove", "cf.add", "cf.remove", "cf.clear", "dv.add", "dv.remove", "banded.set", "outline.set",
                "sheetTable.add", "sheetTable.remove", "sheetTable.update",
                "drawing.add", "drawing.remove", "drawing.transform", "drawing.transform.batch", "drawing.anchor", "drawing.payload.update", "drawing.zorder", "drawing.zorder.restore",
                "pivot.add", "pivot.remove", "pivot.update", "pivot.refresh", "pivot.drilldown.add", "pivot.drilldown.remove",
                "sparkline.add", "sparkline.remove", "sparkline.update", "sparkline.group.add", "sparkline.group.remove", "sparkline.group.replace",
                "table.add", "table.remove", "name.set", "name.remove",
                "print.pageSetup.set", "print.area.set", "print.area.clear", "print.pageBreak.set", "print.pageBreak.remove", "print.pageBreaks.clear", "print.document.replace"
                , "query.definition.replace", "query.load.range", "query.load.sheet-table", "query.load.pivot-source",
                "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted", "cells.shifted", "cells.shifted.restore", "rows.permuted", "sheet.rename"
        ), Set.copyOf(registry.acceptedIds()));
    }

    @Test
    void everyKnownNonAcceptedMutationHasAServerOwnedReason() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        assertEquals(Set.of(
                "automation.recording.changed",
                "sheet.add", "sheet.remove", "sheet.duplicated", "sheet.restore",
                "hyperlink.set", "hyperlink.remove",
                "query.load.workbook-table",
                "workbook.restore"
        ), registry.unavailableReasons().keySet());
        assertEquals("Requires identity remapping for scoped names, object payloads, print state and source relationships.", registry.unavailableReasons().get("sheet.duplicated"));
    }

    @Test
    void commenterMayCommitReviewMutationButCannotWriteCells() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","cells":{},"notes":[],"commentThreads":[]}]}
                """);
        var note = new OperationMutation("note.set", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","row":2,"column":3,"note":{"id":"n-1","author":"guest","text":"Review","createdAt":"2026-08-23T00:00:00Z","visible":true}}
                """));

        var prepared = registry.prepare(snapshot, note, WorkbookAclRole.COMMENTER);
        assertEquals(2, prepared.affectedRanges().get(0).startRow());
        var next = prepared.descriptor().apply(snapshot, note);
        assertEquals("n-1", next.path("sheets").get(0).path("notes").get(0).path("note").path("id").asText());

        var cell = new OperationMutation("cell.set", "sheet-1", mapper.readTree("""
                {"row":0,"column":0,"value":{"value":"no"}}
                """));
        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, cell, WorkbookAclRole.COMMENTER));
        assertEquals("FORBIDDEN", error.code());
    }

    @Test
    void lockedRangeIsResolvedFromSnapshotAndCannotBeBypassedByClientPayload() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","cells":{},"protectionRules":[
                  {"id":"lock-a1","scope":"range","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"locked":true,"allow":{}}
                ]}]}
                """);
        var mutation = new OperationMutation("cell.set", "sheet-1", mapper.readTree("""
                {"row":0,"column":0,"value":{"value":42}}
                """));

        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));
        assertEquals("FORBIDDEN", error.code());
    }

    @Test
    void sheetMetadataMutationsUseCanonicalCollectionsRatherThanClientRanges() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":20,"columnCount":10,"cells":{},"conditionalFormats":[],"dataValidations":[],"sheetTables":[]}],"definedNameModels":[],"printDocuments":[]}
                """);
        var filter = new OperationMutation("filter.set", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","filter":{"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":2},"criteria":{}}}
                """));
        var conditionalFormat = new OperationMutation("cf.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","rule":{"id":"cf-1","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":1,"endRow":3,"startColumn":0,"endColumn":0}],"type":"highlight"}}
                """));
        var sheetTable = new OperationMutation("sheetTable.add", "sheet-1", mapper.readTree("""
                {"id":"table-1","sheetId":"sheet-1","name":"Sales","range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":2},"hasHeaderRow":true,"hasTotalRow":false,"columns":[]}
                """));
        var hideRow = new OperationMutation("row.hidden", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","index":2}
                """));

        JsonNode current = snapshot;
        for (OperationMutation mutation : List.of(filter, conditionalFormat, sheetTable, hideRow)) {
            var prepared = registry.prepare(current, mutation, WorkbookAclRole.EDITOR);
            current = prepared.descriptor().apply(current, mutation);
        }

        var sheet = current.path("sheets").get(0);
        assertEquals(4, sheet.path("filter").path("range").path("endRow").asInt());
        assertEquals("cf-1", sheet.path("conditionalFormats").get(0).path("id").asText());
        assertEquals("table-1", sheet.path("sheetTables").get(0).path("id").asText());
        assertEquals(2, sheet.path("hiddenRows").get(0).asInt());
    }

    @Test
    void drawingReducerKeepsObjectAndPayloadCollectionsInOneAtomicMutation() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":10,"cells":{},"drawings":[],"drawingPayloads":{}}]}
                """);
        OperationMutation add = new OperationMutation("drawing.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","drawing":{"id":"draw-1","sheetId":"sheet-1","kind":"image","payloadId":"payload-1","anchor":{"kind":"absolute"},"transform":{"x":1,"y":2,"width":30,"height":40,"rotation":0},"zIndex":1},"payload":{"kind":"image","src":"data:image/png;base64,AA==","altText":"Logo"}}
                """));
        var prepared = registry.prepare(snapshot, add, WorkbookAclRole.EDITOR);
        JsonNode current = prepared.descriptor().apply(snapshot, add);
        assertEquals("draw-1", current.path("sheets").get(0).path("drawings").get(0).path("id").asText());
        assertEquals("image", current.path("sheets").get(0).path("drawingPayloads").path("payload-1").path("kind").asText());

        OperationMutation update = new OperationMutation("drawing.payload.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","payloadId":"payload-1","before":{"kind":"image","src":"data:image/png;base64,AA==","altText":"Logo"},"after":{"kind":"image","src":"data:image/png;base64,AA==","altText":"Updated"}}
                """));
        current = registry.prepare(current, update, WorkbookAclRole.EDITOR).descriptor().apply(current, update);
        assertEquals("Updated", current.path("sheets").get(0).path("drawingPayloads").path("payload-1").path("altText").asText());
    }

    @Test
    void printAndQueryRangeReducersPersistOnlyCanonicalDomainState() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"schema":"WorkbookSnapshot","unitId":"book-1","name":"Book","sheets":[{"id":"sheet-1","rowCount":20,"columnCount":10,"cells":{}}],"printDocuments":[],"queryDefinitions":[]}
                """);
        OperationMutation print = new OperationMutation("print.document.replace", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","document":{"schema":"PrintDocument","unitId":"book-1","sheetId":"sheet-1","pageSetup":{"paperSize":"a4","orientation":"portrait","margins":{"top":72,"right":72,"bottom":72,"left":72,"header":36,"footer":36},"scale":100,"printGridlines":false,"printHeadings":false,"centerHorizontally":false,"centerVertically":false},"printAreas":[{"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":2}}],"pageBreaks":[]}}
                """));
        JsonNode current = registry.prepare(snapshot, print, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, print);
        assertEquals("a4", current.path("printDocuments").get(0).path("pageSetup").path("paperSize").asText());

        OperationMutation load = new OperationMutation("query.load.range", "sheet-1", mapper.readTree("""
                {"kind":"cells","queryId":"query-1","queryDefinition":{"schema":"QueryDefinition","id":"query-1","name":"Inline","connectorId":"json","connectorConfig":{"data":[]},"steps":[],"sourceRevision":0},"target":{"kind":"range","sheetId":"sheet-1"},"clearRange":{"sheetId":"sheet-1","startRow":0,"endRow":2,"startColumn":0,"endColumn":1},"values":[[{"value":"Name"},{"value":"Amount"}],[{"value":"East"},{"value":42}]]}
                """));
        current = registry.prepare(current, load, WorkbookAclRole.EDITOR).descriptor().apply(current, load);
        assertEquals("query-1", current.path("queryDefinitions").get(0).path("id").asText());
        assertEquals("East", current.path("sheets").get(0).path("cells").path("1").path("0").path("value").asText());
        assertEquals(42, current.path("sheets").get(0).path("cells").path("1").path("1").path("value").asInt());
    }

    @Test
    void pivotSparklineAndDrillDownReducersPreserveOnlyDomainDefinitionsAndDerivedDetailCells() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sales","rowCount":20,"columnCount":10,"cells":{"0":{"0":{"value":"Region"},"1":{"value":"Amount"}},"1":{"0":{"value":"East"},"1":{"value":42}}},"pivots":[],"sparklines":[],"sparklineGroups":[]}]}
                """);
        OperationMutation pivot = new OperationMutation("pivot.add", "sheet-1", mapper.readTree("""
                {"id":"pivot-1","sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1},"layout":{"rows":[],"columns":[],"filters":[],"values":[],"showSubtotals":true,"showGrandTotals":true}}
                """));
        JsonNode current = registry.prepare(snapshot, pivot, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, pivot);
        assertEquals("pivot-1", current.path("sheets").get(0).path("pivots").get(0).path("id").asText());

        OperationMutation sparkline = new OperationMutation("sparkline.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sparkline":{"id":"spark-1","sheetId":"sheet-1","anchor":{"row":3,"column":0},"sourceRange":{"sheetId":"sheet-1","startRow":1,"endRow":1,"startColumn":1,"endColumn":1},"type":"line","color":"#2563eb"}}
                """));
        current = registry.prepare(current, sparkline, WorkbookAclRole.EDITOR).descriptor().apply(current, sparkline);
        assertEquals("spark-1", current.path("sheets").get(0).path("sparklines").get(0).path("id").asText());

        OperationMutation drillDown = new OperationMutation("pivot.drilldown.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","pivotId":"pivot-1","label":"East","sourceRowPaths":[{"sheetId":"sheet-1","row":1}],"targetSheetId":"detail-1","targetAnchor":{"row":0,"column":0}}
                """));
        current = registry.prepare(current, drillDown, WorkbookAclRole.EDITOR).descriptor().apply(current, drillDown);
        assertEquals("detail-1", current.path("sheets").get(1).path("id").asText());
        assertEquals("Region", current.path("sheets").get(1).path("cells").path("0").path("0").path("value").asText());
        assertEquals("East", current.path("sheets").get(1).path("cells").path("1").path("0").path("value").asText());
    }

    @Test
    void sparklineGroupStateIsAppliedAsOneConsistentMutation() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":10,"cells":{},"sparklines":[
                  {"id":"s1","sheetId":"sheet-1","anchor":{"row":1,"column":0},"sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"type":"line","color":"#1"},
                  {"id":"s2","sheetId":"sheet-1","anchor":{"row":2,"column":0},"sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"type":"line","color":"#1"}
                ],"sparklineGroups":[]}]}
                """);
        OperationMutation group = new OperationMutation("sparkline.group.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","groupIds":["g1"],"groups":[{"index":0,"group":{"id":"g1","sheetId":"sheet-1","type":"line","sparklineIds":["s1","s2"],"showAxis":true}}],"members":[{"sparklineId":"s1","type":"line","groupId":"g1","showAxis":true},{"sparklineId":"s2","type":"line","groupId":"g1","showAxis":true}]}
                """));

        JsonNode current = registry.prepare(snapshot, group, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, group);
        assertEquals("g1", current.path("sheets").get(0).path("sparklineGroups").get(0).path("id").asText());
        assertEquals("g1", current.path("sheets").get(0).path("sparklines").get(0).path("groupId").asText());
        assertEquals(true, current.path("sheets").get(0).path("sparklines").get(1).path("showAxis").asBoolean());
    }

    @Test
    void structuralRowMutationMovesCellsMetadataAndParsedFormulaReferencesTogether() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"definedNames":{"Sales":"=Sheet1!A2"},"definedNameModels":[{"name":"Sales","formula":"=Sheet1!A2","scope":"workbook"}],"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=A2"}},"1":{"0":{"value":10}}},"freeze":{"xSplit":0,"ySplit":1,"startRow":1,"startColumn":0},"hiddenRows":[1],"rowHeights":{"1":33},"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"notes":[],"commentThreads":[],"spillRanges":[],"protectionRules":[]},
                  {"id":"sheet-2","name":"Other","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=Sheet1!A2"}}},"freeze":{"xSplit":0,"ySplit":0,"startRow":0,"startColumn":0}}
                ]}
                """);
        OperationMutation insert = new OperationMutation("rows.inserted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","at":1,"count":1}
                """));

        JsonNode current = registry.prepare(snapshot, insert, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, insert);
        assertEquals(6, current.path("sheets").get(0).path("rowCount").asInt());
        assertEquals(10, current.path("sheets").get(0).path("cells").path("2").path("0").path("value").asInt());
        assertEquals("=A3", current.path("sheets").get(0).path("cells").path("0").path("0").path("formula").asText());
        assertEquals("=Sheet1!A3", current.path("sheets").get(1).path("cells").path("0").path("0").path("formula").asText());
        assertEquals("=Sheet1!A3", current.path("definedNames").path("Sales").asText());
        assertEquals(2, current.path("sheets").get(0).path("hiddenRows").get(0).asInt());
    }

    @Test
    void sheetRenameUsesParsedReferencesAndDoesNotTouchStringLiterals() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","name":"Old Name","rowCount":5,"columnCount":3,"cells":{},"freeze":{"xSplit":0,"ySplit":0,"startRow":0,"startColumn":0}},
                  {"id":"sheet-2","name":"Other","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"='Old Name'!A1+\\\"Old Name!A1\\\""}}},"freeze":{"xSplit":0,"ySplit":0,"startRow":0,"startColumn":0}}
                ],"definedNames":{"Total":"='Old Name'!A1"},"definedNameModels":[{"name":"Total","formula":"='Old Name'!A1","scope":"workbook"}]}
                """);
        OperationMutation rename = new OperationMutation("sheet.rename", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","name":"New Name"}
                """));

        JsonNode current = registry.prepare(snapshot, rename, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, rename);
        assertEquals("New Name", current.path("sheets").get(0).path("name").asText());
        assertEquals("='New Name'!A1+\"Old Name!A1\"", current.path("sheets").get(1).path("cells").path("0").path("0").path("formula").asText());
        assertEquals("='New Name'!A1", current.path("definedNames").path("Total").asText());
    }

    @Test
    void boundedShiftAndRowPermutationHaveDeterministicInverseFriendlySnapshots() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=A1"}},"1":{"0":{"value":"drop"}}},"freeze":{"xSplit":0,"ySplit":0,"startRow":0,"startColumn":0},"notes":[{"row":0,"column":0,"note":{"id":"n1"}}],"commentThreads":[],"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                """);
        OperationMutation shift = new OperationMutation("cells.shifted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"direction":"down"}
                """));
        JsonNode current = registry.prepare(snapshot, shift, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, shift);
        assertEquals("=A2", current.path("sheets").get(0).path("cells").path("1").path("0").path("formula").asText());
        assertEquals(1, current.path("sheets").get(0).path("notes").get(0).path("row").asInt());

        OperationMutation restore = new OperationMutation("cells.shifted.restore", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"direction":"down","cells":[{"row":0,"column":0,"cell":{"value":null,"formula":"=A1"}},{"row":1,"column":0,"cell":{"value":"drop"}}]}
                """));
        current = registry.prepare(current, restore, WorkbookAclRole.EDITOR).descriptor().apply(current, restore);
        assertEquals("=A1", current.path("sheets").get(0).path("cells").path("0").path("0").path("formula").asText());
        assertEquals("drop", current.path("sheets").get(0).path("cells").path("1").path("0").path("value").asText());

        OperationMutation permutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceRows":[1,0]}
                """));
        current = registry.prepare(current, permutation, WorkbookAclRole.EDITOR).descriptor().apply(current, permutation);
        assertEquals("drop", current.path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());
        assertEquals("=A1", current.path("sheets").get(0).path("cells").path("1").path("0").path("formula").asText());
    }
}
