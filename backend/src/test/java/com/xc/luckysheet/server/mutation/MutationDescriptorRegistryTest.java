package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
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
    void rangePasteAppliesCanonicalSnapshotAndClearsABoundedSource() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":20,"cells":{"0":{"0":{"value":"move"}}}}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":1,"column":1},"sourceExtent":{"rows":1,"columns":1},
                 "transfer":"move","clearSource":true,"sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},
                 "spec":{"content":"all","formatting":"all","metadata":{"commentsNotes":true,"validation":true,"columnWidths":false,"conditionalFormats":true,"hyperlinks":true},"operation":"none","skipBlanks":false,"transpose":false,"link":false},
                 "snapshot":{"cells":[{"row":0,"column":0},{"row":1,"column":1,"value":{"value":"move"}}]}}
                """));

        var prepared = registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR);
        var next = prepared.descriptor().apply(snapshot, mutation);

        assertEquals(2, prepared.affectedRanges().size());
        assertEquals(true, next.path("sheets").get(0).path("cells").path("0").isMissingNode());
        assertEquals("move", next.path("sheets").get(0).path("cells").path("1").path("1").path("value").asText());
    }

    @Test
    void rangePasteRejectsAnOversizedSourceBeforeApplyingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":20,"cells":{"0":{"0":{"value":"keep"}}}}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":1,"column":1},"sourceExtent":{"rows":1,"columns":1},
                 "transfer":"move","clearSource":true,"sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":100000,"startColumn":0,"endColumn":0},
                 "spec":{"content":"all","formatting":"all","metadata":{"commentsNotes":true,"validation":true,"columnWidths":false,"conditionalFormats":true,"hyperlinks":true},"operation":"none","skipBlanks":false,"transpose":false,"link":false},
                 "snapshot":{"cells":[{"row":0,"column":0},{"row":1,"column":1,"value":{"value":"move"}}]}}
                """));

        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));

        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals("keep", snapshot.path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());
    }

    @Test
    void rangePasteRejectsTheRemovedMatrixModeContract() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("{\"sheets\":[{\"id\":\"sheet-1\",\"rowCount\":10,\"columnCount\":10,\"cells\":{}}]}");
        var legacy = new OperationMutation("range.paste", "sheet-1", mapper.readTree("{\"startRow\":0,\"startColumn\":0,\"values\":[[{\"value\":1}]]}"));
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, legacy, WorkbookAclRole.EDITOR));
    }

    @Test
    void rangePasteAcceptsMetadataWhoseEveryOwnedRangeIsInsideTheTarget() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{},"dataValidations":[]}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":0,"column":0},"sourceExtent":{"rows":2,"columns":1},
                 "transfer":"copy","clearSource":false,
                 "spec":{"content":"all","formatting":"all","metadata":{"commentsNotes":true,"validation":true,"columnWidths":false,"conditionalFormats":true,"hyperlinks":true},"operation":"none","skipBlanks":false,"transpose":false,"link":false},
                 "snapshot":{"cells":[],"validations":[{"id":"dv-1","ranges":[
                   {"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},
                   {"sheetId":"sheet-1","startRow":1,"endRow":1,"startColumn":0,"endColumn":0}
                 ]}]}}
                """));

        var prepared = registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR);
        var updated = prepared.descriptor().apply(snapshot, mutation);

        assertEquals("dv-1", updated.path("sheets").get(0).path("dataValidations").get(0).path("id").asText());
    }

    @Test
    void rangePasteRejectsMetadataRuleThatMixesTargetAndUnrelatedRanges() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{},"dataValidations":[],"protectionRules":[
                  {"id":"lock-e5","scope":"range","range":{"sheetId":"sheet-1","startRow":4,"endRow":4,"startColumn":4,"endColumn":4},"locked":true,"allow":{}}
                ]}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":0,"column":0},"sourceExtent":{"rows":1,"columns":1},
                 "transfer":"copy","clearSource":false,
                 "spec":{"content":"all","formatting":"all","metadata":{"commentsNotes":true,"validation":true,"columnWidths":false,"conditionalFormats":true,"hyperlinks":true},"operation":"none","skipBlanks":false,"transpose":false,"link":false},
                 "snapshot":{"cells":[],"validations":[{"id":"dv-attack","ranges":[
                   {"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},
                   {"sheetId":"sheet-1","startRow":4,"endRow":4,"startColumn":4,"endColumn":4}
                 ]}]}}
                """));

        var prepared = registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR);
        ServiceException error = assertThrows(ServiceException.class, () -> prepared.descriptor().apply(snapshot, mutation));

        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals(0, snapshot.path("sheets").get(0).path("dataValidations").size());
    }

    @Test
    void rangePasteReportsAndProtectsTheWholeColumnForWidthChanges() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{},"columnWidthsPx":{},"protectionRules":[
                  {"id":"lock-a10","scope":"range","range":{"sheetId":"sheet-1","startRow":9,"endRow":9,"startColumn":0,"endColumn":0},"locked":true,"allow":{}}
                ]}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":0,"column":0},"sourceExtent":{"rows":1,"columns":1},
                 "transfer":"copy","clearSource":false,
                 "spec":{"content":"all","formatting":"all","metadata":{"commentsNotes":true,"validation":true,"columnWidths":true,"conditionalFormats":true,"hyperlinks":true},"operation":"none","skipBlanks":false,"transpose":false,"link":false},
                 "snapshot":{"cells":[],"columnWidths":[{"column":0,"widthPx":120}]}}
                """));

        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));
        assertEquals("FORBIDDEN", error.code());

        var owner = registry.prepare(snapshot, mutation, WorkbookAclRole.OWNER);
        assertEquals(1_048_575, owner.affectedRanges().get(1).endRow());
        assertEquals(0, owner.affectedRanges().get(1).startColumn());
    }

    @Test
    void internalRestoreCannotBeSubmittedByClient() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        assertThrows(ServiceException.class, () -> registry.require("workbook.restore", false));
    }

    @Test
    void knownRemoteStructureMutationHasAnExplicitServerReducer() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        assertEquals("sheet.duplicated", registry.require("sheet.duplicated", false).id());
        assertEquals(true, registry.ids().contains("sheet.duplicated"));
        assertEquals(true, registry.acceptedIds().contains("sheet.duplicated"));
    }

    @Test
    void acceptedMutationSurfaceIsExplicitAndAllOtherKnownMutationsRemainFailClosed() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        assertEquals(Set.of(
                "cell.set", "cell.restore", "cell.editor.set", "cellTemplate.set", "cellTemplate.remove", "range.set", "range.paste", "range.clear", "range.clear.restore",
                "style.set", "merge.set", "merge.remove", "freeze.set", "row.resize", "column.resize", "column.defaultWidth.resize", "columns.visibility", "view.set", "sheet.hidden", "sheet.unhidden", "sheet.tabColor",
                "note.set", "note.remove", "note.visibility", "comment.add", "comment.reply", "comment.reply.remove", "comment.resolve", "comment.remove",
                "sheet.protect.set", "sheet.protect.remove", "workbook.renamed",
                "sheet.add", "sheet.remove", "sheet.rename", "sheet.duplicated", "sheet.restore", "hyperlink.set", "hyperlink.remove",
                "sheet.reordered",
                "row.hidden", "row.unhidden", "rows.unhidden.all", "rows.hidden.restore",
                "column.hidden", "column.unhidden", "columns.unhidden.all", "columns.hidden.restore",
                "autoFilter.set", "autoFilter.remove", "cf.add", "cf.remove", "cf.clear", "dv.add", "dv.remove", "banded.set", "outline.set",
                "sheetTable.add", "sheetTable.remove", "sheetTable.update", "sheetTable.autoFilter.set", "tableSheet.update", "ganttSheet.update", "reportSheet.update",
                "drawing.add", "drawing.remove", "drawing.transform", "drawing.transform.batch", "drawing.anchor", "drawing.payload.update", "drawing.zorder", "drawing.zorder.restore",
                "pivot.add", "pivot.remove", "pivot.update", "pivot.refresh", "pivot.drilldown.add", "pivot.drilldown.remove",
                "sparkline.add", "sparkline.remove", "sparkline.update", "sparkline.group.add", "sparkline.group.remove", "sparkline.group.replace",
                "table.add", "table.remove", "name.set", "name.remove",
                "print.pageSetup.set", "print.area.set", "print.area.clear", "print.pageBreak.set", "print.pageBreak.remove", "print.pageBreaks.clear", "print.document.replace"
                , "query.definition.replace", "query.load.range", "query.load.sheet-table", "query.load.pivot-source",
                "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted", "cells.inserted", "cells.deleted", "cells.inserted.restore", "cells.deleted.restore", "rows.permuted",
                "dataSource.add", "dataSource.update", "dataSource.remove", "dataRegion.add", "dataRegion.remove"
        ), Set.copyOf(registry.acceptedIds()));
    }

    @Test
    void tableSheetUpdateUsesTheBoundTableAndWholeSheetRange() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {
                  "dataModel":{"tables":[{"id":"table-1","fields":[{"id":"name"},{"id":"amount"}]}]},
                  "sheets":[{"id":"sheet-1","kind":"table-sheet","rowCount":20,"columnCount":5,"cells":{},
                    "tableSheet":{"viewId":"table-1","columns":[{"fieldId":"name","caption":"Name"}],"grouping":[]}}
                ]}
                """);
        var update = new OperationMutation("tableSheet.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","definition":{"viewId":"table-1","columns":[{"fieldId":"amount","caption":"Amount","widthPx":120}],"grouping":[],"sortState":[{"fieldId":"amount","direction":"desc"}]}}
                """));

        var prepared = registry.prepare(snapshot, update, WorkbookAclRole.EDITOR);
        assertEquals(1, prepared.affectedRanges().size());
        assertEquals(0, prepared.affectedRanges().getFirst().startRow());
        assertEquals(19, prepared.affectedRanges().getFirst().endRow());
        assertEquals(0, prepared.affectedRanges().getFirst().startColumn());
        assertEquals(4, prepared.affectedRanges().getFirst().endColumn());
        var updated = registry.applyPublicMutations(snapshot, List.of(update));
        assertEquals("amount", updated.path("sheets").get(0).path("tableSheet").path("columns").get(0).path("fieldId").asText());

        var invalid = new OperationMutation("tableSheet.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","definition":{"viewId":"table-1","columns":[{"fieldId":"missing","caption":"Missing"}],"grouping":[]}}
                """));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(invalid)));
        assertEquals("name", snapshot.path("sheets").get(0).path("tableSheet").path("columns").get(0).path("fieldId").asText());
    }

    @Test
    void ganttSheetUpdateUsesTheBoundTableAndWholeSheetRange() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {
                  "dataModel":{"tables":[{"id":"tasks","fields":[{"id":"id"},{"id":"title"},{"id":"start"},{"id":"end"},{"id":"progress"},{"id":"parent"},{"id":"deps"}]}]},
                  "sheets":[{"id":"sheet-1","kind":"gantt-sheet","rowCount":20,"columnCount":8,"cells":{},
                    "ganttSheet":{"viewId":"tasks","fieldMap":{"id":"id","title":"title","start":"start","end":"end","progress":"progress","parentId":"parent","dependencies":"deps"},"calendar":{"workingDays":[1,2,3,4,5],"dayStartHour":9,"dayEndHour":18},"timeline":{"unit":"week"},"dependencyStyle":{"color":"#64748b","width":1}}}
                ]}
                """);
        var update = new OperationMutation("ganttSheet.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","definition":{"viewId":"tasks","fieldMap":{"id":"id","title":"title","start":"start","end":"end","progress":"progress","parentId":"parent","dependencies":"deps"},"calendar":{"workingDays":[1,2,3,4,5],"dayStartHour":8,"dayEndHour":17},"timeline":{"unit":"day"},"dependencyStyle":{"color":"#334155","width":2}}}
                """));
        var prepared = registry.prepare(snapshot, update, WorkbookAclRole.EDITOR);
        assertEquals(1, prepared.affectedRanges().size());
        assertEquals(19, prepared.affectedRanges().getFirst().endRow());
        var updated = registry.applyPublicMutations(snapshot, List.of(update));
        assertEquals("day", updated.path("sheets").get(0).path("ganttSheet").path("timeline").path("unit").asText());
        var invalid = new OperationMutation("ganttSheet.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","definition":{"viewId":"tasks","fieldMap":{"id":"missing","title":"title","start":"start","end":"end","progress":"progress"},"calendar":{"workingDays":[1],"dayStartHour":9,"dayEndHour":18},"timeline":{"unit":"week"},"dependencyStyle":{"color":"#64748b","width":1}}}
                """));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(invalid)));
    }

    @Test
    void reportSheetUpdateValidatesTemplateBindingsAndWholeSheetRange() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"dataModel":{"tables":[{"id":"tasks","fields":[{"id":"title"}]}]},"sheets":[
                  {"id":"template","kind":"worksheet","rowCount":20,"columnCount":8,"cells":{}},
                  {"id":"report","kind":"report-sheet","rowCount":20,"columnCount":8,"cells":{},"reportSheet":{"templateSheetId":"template","tableId":"tasks","bindings":[],"pagination":{"enabled":true,"rowsPerPage":10,"repeatHeaderRows":[0]},"renderMode":"design","layout":{"orientation":"portrait","marginTopPx":24,"marginRightPx":24,"marginBottomPx":24,"marginLeftPx":24},"dataEntry":[]}}
                ]}
                """);
        var update = new OperationMutation("reportSheet.update", "report", mapper.readTree("""
                {"sheetId":"report","definition":{"templateSheetId":"template","tableId":"tasks","bindings":[{"cell":{"row":1,"column":0},"expression":"title","kind":"field","direction":"vertical","fill":"down"}],"pagination":{"enabled":true,"rowsPerPage":5,"repeatHeaderRows":[0]},"renderMode":"preview","layout":{"orientation":"landscape","marginTopPx":12,"marginRightPx":12,"marginBottomPx":12,"marginLeftPx":12},"dataEntry":[{"fieldId":"title","writable":true}]}}
                """));
        var prepared = registry.prepare(snapshot, update, WorkbookAclRole.EDITOR);
        assertEquals(1, prepared.affectedRanges().size());
        assertEquals(19, prepared.affectedRanges().getFirst().endRow());
        var updated = registry.applyPublicMutations(snapshot, List.of(update));
        assertEquals("preview", updated.path("sheets").get(1).path("reportSheet").path("renderMode").asText());
        var invalid = new OperationMutation("reportSheet.update", "report", mapper.readTree("""
                {"sheetId":"report","definition":{"templateSheetId":"template","tableId":"tasks","bindings":[{"cell":{"row":1,"column":0},"expression":"missing","kind":"field"}],"pagination":{"enabled":true,"rowsPerPage":5},"renderMode":"design","layout":{"orientation":"portrait","marginTopPx":0,"marginRightPx":0,"marginBottomPx":0,"marginLeftPx":0},"dataEntry":[]}}
                """));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(invalid)));
    }

    @Test
    void everyKnownNonAcceptedMutationHasAServerOwnedReason() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        assertEquals(Set.of(
                "automation.recording.changed",
                "pivot.chart.create",
                "query.load.workbook-table",
                "workbook.restore"
        ), registry.unavailableReasons().keySet());
        assertEquals("Recorder state is transient session state and must not enter workbook history.", registry.unavailableReasons().get("automation.recording.changed"));
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
    void freezeSetAcceptsCanonicalPaneStates() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","pane":{"kind":"none"}}]}
                """);

        for (String pane : List.of(
                "{\"kind\":\"frozen\",\"state\":\"frozen\",\"xSplit\":1,\"ySplit\":0,\"startRow\":0,\"startColumn\":1}",
                "{\"kind\":\"frozen\",\"state\":\"frozenSplit\",\"xSplit\":1,\"ySplit\":1,\"startRow\":1,\"startColumn\":1}",
                "{\"kind\":\"split\",\"state\":\"split\",\"xSplit\":20.5,\"ySplit\":10,\"startRow\":1,\"startColumn\":2}")) {
            OperationMutation mutation = new OperationMutation("freeze.set", "sheet-1", mapper.readTree("{\"pane\":" + pane + "}"));
            snapshot = registry.applyPublicMutations(snapshot, List.of(mutation));
            assertEquals(mapper.readTree(pane), snapshot.path("sheets").get(0).path("pane"));
        }
    }

    @Test
    void freezeSetRejectsPaneStateThatWouldPoisonCanonicalSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","pane":{"kind":"none"}}]}
                """);

        for (String pane : List.of(
                "{\"kind\":\"frozen\",\"xSplit\":1,\"ySplit\":0,\"startRow\":0,\"startColumn\":1}",
                "{\"kind\":\"frozen\",\"state\":\"split\",\"xSplit\":1,\"ySplit\":0,\"startRow\":0,\"startColumn\":1}",
                "{\"kind\":\"split\",\"state\":\"frozen\",\"xSplit\":1,\"ySplit\":1,\"startRow\":1,\"startColumn\":1}")) {
            OperationMutation mutation = new OperationMutation("freeze.set", "sheet-1", mapper.readTree("{\"pane\":" + pane + "}"));
            ServiceException error = assertThrows(ServiceException.class,
                    () -> registry.applyPublicMutations(snapshot, List.of(mutation)));
            assertEquals("VALIDATION_ERROR", error.code());
        }
    }

    @Test
    void rowPermutationChecksProtectedMetadataAcrossEveryColumnItRemaps() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":5,"cells":{},"protectionRules":[
                  {"id":"lock-outside-grid","scope":"range","range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":500,"endColumn":500},"locked":true,"allow":{}}
                ]}]}
                """);
        var mutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":9,"startColumn":0,"endColumn":0},"sourceRows":[5,6,7,8,9,0,1,2,3,4]}
                """));

        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));
        assertEquals("FORBIDDEN", error.code());

        var owner = registry.prepare(snapshot, mutation, WorkbookAclRole.OWNER);
        assertEquals(0, owner.affectedRanges().getFirst().startColumn());
        assertEquals(16_383, owner.affectedRanges().getFirst().endColumn());
        var updated = owner.descriptor().apply(snapshot, mutation);
        assertEquals(5, updated.path("sheets").get(0).path("protectionRules").get(0).path("range").path("startRow").asInt());
    }

    @Test
    void sheetMetadataMutationsUseCanonicalCollectionsRatherThanClientRanges() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":20,"columnCount":10,"cells":{},"conditionalFormats":[],"dataValidations":[],"sheetTables":[]}],"definedNameModels":[],"printDocuments":[]}
                """);
        var filter = new OperationMutation("autoFilter.set", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","autoFilter":{"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":2},"columns":{}}}
                """));
        var conditionalFormat = new OperationMutation("cf.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","rule":{"id":"cf-1","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":1,"endRow":3,"startColumn":0,"endColumn":0}],"type":"highlight"}}
                """));
        var sheetTable = new OperationMutation("sheetTable.add", "sheet-1", mapper.readTree("""
                {"id":"table-1","sheetId":"sheet-1","name":"Sales","range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":2},"hasHeaderRow":true,"hasTotalRow":false,"showBandedRows":true,"showBandedColumns":false,"showFirstColumn":false,"showLastColumn":false,"showFilterButton":true,"autoExpand":"both","columns":[{"id":"c1","name":"A"},{"id":"c2","name":"B"},{"id":"c3","name":"C"}]}
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
        assertEquals(4, sheet.path("autoFilter").path("range").path("endRow").asInt());
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
    void drawingReducerRejectsCameraRangesThatCanExhaustTheRenderer() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":1000,"columnCount":1000,"cells":{},"drawings":[],"drawingPayloads":{}}]}
                """);
        OperationMutation bounded = new OperationMutation("drawing.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","drawing":{"id":"bounded-camera","sheetId":"sheet-1","kind":"camera","payloadId":"bounded-payload","anchor":{"kind":"absolute"},"transform":{"x":1,"y":2,"width":30,"height":40,"rotation":0},"zIndex":1},"payload":{"kind":"camera","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":99,"startColumn":0,"endColumn":99},"refreshPolicy":"live"}}
                """));
        registry.prepare(snapshot, bounded, WorkbookAclRole.EDITOR);

        OperationMutation add = new OperationMutation("drawing.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","drawing":{"id":"camera-1","sheetId":"sheet-1","kind":"camera","payloadId":"camera-payload","anchor":{"kind":"absolute"},"transform":{"x":1,"y":2,"width":30,"height":40,"rotation":0},"zIndex":1},"payload":{"kind":"camera","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":999,"startColumn":0,"endColumn":999},"refreshPolicy":"live"}}
                """));

        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, add, WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", error.code());
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
                {"schema":"PivotDefinition","id":"pivot-1","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[{"fieldId":"sheet:sheet-1:column:0:range:0","name":"Region","dataType":"text","ordinal":0},{"fieldId":"sheet:sheet-1:column:1:range:0","name":"Amount","dataType":"number","ordinal":1}]},"layout":{"rows":[{"fieldId":"sheet:sheet-1:column:0:range:0","subtotal":{"mode":"automatic"}}],"columns":[],"filters":[{"kind":"manual","family":"manual","fieldId":"sheet:sheet-1:column:0:range:0","scope":"report","mode":"all","memberKeys":[]}],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[{"valueId":"value:amount","fieldId":"sheet:sheet-1:column:1:range:0","summarizeBy":"sum"}],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                """));
        JsonNode current = registry.prepare(snapshot, pivot, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, pivot);
        assertEquals("pivot-1", current.path("sheets").get(0).path("pivots").get(0).path("id").asText());

        ObjectNode highCardinality = (ObjectNode) pivot.params().deepCopy();
        ArrayNode members = mapper.createArrayNode();
        for (int index = 0; index < 10_001; index++) members.add("Member " + index);
        ((ObjectNode) highCardinality.path("fieldCatalog").path("fields").get(0)).set("values", members);
        OperationMutation highCardinalityMutation = new OperationMutation("pivot.add", "sheet-1", highCardinality);
        JsonNode highCardinalitySnapshot = registry.prepare(snapshot, highCardinalityMutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, highCardinalityMutation);
        assertEquals(10_001, highCardinalitySnapshot.path("sheets").get(0).path("pivots").get(0).path("fieldCatalog").path("fields").get(0).path("values").size());

        ObjectNode malformedValues = (ObjectNode) pivot.params().deepCopy();
        ((ObjectNode) malformedValues.path("fieldCatalog").path("fields").get(0)).set("values", mapper.createObjectNode());
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, new OperationMutation("pivot.add", "sheet-1", malformedValues), WorkbookAclRole.EDITOR));

        OperationMutation sparkline = new OperationMutation("sparkline.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sparkline":{"id":"spark-1","sheetId":"sheet-1","anchor":{"row":3,"column":0},"sourceRange":{"sheetId":"sheet-1","startRow":1,"endRow":1,"startColumn":1,"endColumn":1},"type":"line","color":"#2563eb"}}
                """));
        current = registry.prepare(current, sparkline, WorkbookAclRole.EDITOR).descriptor().apply(current, sparkline);
        assertEquals("spark-1", current.path("sheets").get(0).path("sparklines").get(0).path("id").asText());

        OperationMutation drillDown = new OperationMutation("pivot.drilldown.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","pivotId":"pivot-1","label":"East","sourceRowPaths":[{"sheetId":"sheet-1","row":1}],"targetSheetId":"detail-1","target":{"row":0,"column":0}}
                """));
        current = registry.prepare(current, drillDown, WorkbookAclRole.EDITOR).descriptor().apply(current, drillDown);
        assertEquals("detail-1", current.path("sheets").get(1).path("id").asText());
        assertEquals("Region", current.path("sheets").get(1).path("cells").path("0").path("0").path("value").asText());
        assertEquals("East", current.path("sheets").get(1).path("cells").path("1").path("0").path("value").asText());
    }

    @Test
    void pivotMutationsRejectLegacyShapeAndRefreshDoesNotPersistRuntimeState() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":10,"cells":{},"pivots":[]}]}
                """);
        OperationMutation legacy = new OperationMutation("pivot.add", "sheet-1", mapper.readTree("""
                {"id":"pivot-legacy","sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1},"layout":{"rows":[],"columns":[],"filters":[],"values":[],"showSubtotals":true,"showGrandTotals":true}}
                """));
        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, legacy, WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", error.code());

        OperationMutation oldLayout = new OperationMutation("pivot.add", "sheet-1", mapper.readTree("""
                {"schema":"PivotDefinition","id":"pivot-old-layout","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[]},"layout":{"rows":[],"columns":[],"filters":[],"values":[],"showGrandTotals":true,"compact":false,"repeatLabels":false},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                """));
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, oldLayout, WorkbookAclRole.EDITOR));

        OperationMutation malformedSubtotal = new OperationMutation("pivot.add", "sheet-1", mapper.readTree("""
                {"schema":"PivotDefinition","id":"pivot-bad-subtotal","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[{"fieldId":"sheet:sheet-1:column:0:range:0","name":"Region","dataType":"text","ordinal":0}]},"layout":{"rows":[{"fieldId":"sheet:sheet-1:column:0:range:0","subtotal":{"mode":"custom","functions":[]}}],"columns":[],"filters":[],"values":[],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                """));
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, malformedSubtotal, WorkbookAclRole.EDITOR));

        OperationMutation add = new OperationMutation("pivot.add", "sheet-1", mapper.readTree("""
                {"schema":"PivotDefinition","id":"pivot-1","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[]},"layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                """));
        JsonNode current = registry.applyPublicMutations(snapshot, List.of(add));
        JsonNode beforeRefresh = current.deepCopy();
        OperationMutation refresh = new OperationMutation("pivot.refresh", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","pivotId":"pivot-1"}
                """));
        JsonNode afterRefresh = registry.applyPublicMutations(current, List.of(refresh));
        assertEquals(beforeRefresh, afterRefresh);
    }

    @Test
    void pivotLayoutOnlyUpdateAcceptsNewCalculatedFieldsWithoutWritingTheFieldCatalog() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sales","rowCount":20,"columnCount":10,"cells":{"0":{"0":{"value":"Region"},"1":{"value":"Amount"}},"1":{"0":{"value":"East"},"1":{"value":42}}},"pivots":[
                  {"schema":"PivotDefinition","id":"pivot-1","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[{"fieldId":"sheet:sheet-1:column:0:range:0","name":"Region","dataType":"text","ordinal":0},{"fieldId":"sheet:sheet-1:column:1:range:0","name":"Amount","dataType":"number","ordinal":1}]},"layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                ]}]}
                """);
        OperationMutation update = new OperationMutation("pivot.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","pivotId":"pivot-1","layout":{"rows":[{"fieldId":"calculated:margin"}],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[{"valueId":"value:calculated:margin","fieldId":"calculated:margin","summarizeBy":"sum"}],"calculatedFields":[{"fieldId":"calculated:margin","name":"Margin","formula":"=amount*1.15"}],"calculatedItems":[{"fieldId":"calculated-item:amount:premium","targetFieldId":"sheet:sheet-1:column:1:range:0","name":"Premium","formula":"=amount*3"}],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"}}
                """));

        JsonNode next = registry.applyPublicMutations(snapshot, List.of(update));
        JsonNode pivot = next.path("sheets").get(0).path("pivots").get(0);
        assertEquals("calculated:margin", pivot.path("layout").path("calculatedFields").get(0).path("fieldId").asText());
        assertEquals("calculated-item:amount:premium", pivot.path("layout").path("calculatedItems").get(0).path("fieldId").asText());
        assertEquals(2, pivot.path("fieldCatalog").path("fields").size());
        assertEquals(false, pivot.path("fieldCatalog").path("fields").toString().contains("calculated:margin"));
    }

    @Test
    void pivotCalculatedDefinitionsRejectCatalogCollisionsAndUnknownTargets() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sales","rowCount":20,"columnCount":10,"cells":{},"pivots":[
                  {"schema":"PivotDefinition","id":"pivot-1","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[{"fieldId":"sheet:sheet-1:column:0:range:0","name":"Region","dataType":"text","ordinal":0},{"fieldId":"sheet:sheet-1:column:1:range:0","name":"Amount","dataType":"number","ordinal":1}]},"layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                ]}]}
                """);
        OperationMutation collision = new OperationMutation("pivot.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","pivotId":"pivot-1","layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"calculatedFields":[{"fieldId":"sheet:sheet-1:column:0:range:0","name":"Shadow","formula":"=1"}],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"}}
                """));
        ServiceException collisionError = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, collision, WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", collisionError.code());

        OperationMutation missingTarget = new OperationMutation("pivot.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","pivotId":"pivot-1","layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"calculatedFields":[],"calculatedItems":[{"fieldId":"calculated-item:missing","targetFieldId":"missing","name":"Missing","formula":"=1"}],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"}}
                """));
        ServiceException targetError = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, missingTarget, WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", targetError.code());
    }

    @Test
    void pivotFiltersPersistCanonicalScopesAndRejectInvalidAxisScope() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sales","rowCount":20,"columnCount":10,
                  "cells":{"0":{"0":{"value":"Region"},"1":{"value":"Amount"}},"1":{"0":{"value":"East"},"1":{"value":42}},"2":{"0":{"value":"West"},"1":{"value":18}}},"pivots":[]}]}
                """);
        String region = "sheet:sheet-1:column:0:range:0";
        String amount = "sheet:sheet-1:column:1:range:0";
        String valueId = "value:amount:sum";
        String countValueId = "value:amount:count";
        ObjectNode pivot = (ObjectNode) mapper.readTree("""
                {"schema":"PivotDefinition","id":"pivot-filter-scope","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":2,"startColumn":0,"endColumn":1}},
                 "target":{"sheetId":"sheet-1","anchor":{"row":5,"column":3}},
                 "fieldCatalog":{"schema":"PivotFieldCatalog","fields":[{"fieldId":"%s","name":"Region","dataType":"text","ordinal":0},{"fieldId":"%s","name":"Amount","dataType":"number","ordinal":1}]},
                 "layout":{"rows":[{"fieldId":"%s"}],"columns":[],"filters":[
                   {"kind":"manual","family":"manual","fieldId":"%s","mode":"all","memberKeys":[]},
                   {"kind":"condition","family":"label","fieldId":"%s","scope":"report","operator":"begins-with","value":"E"},
                   {"kind":"top-items","family":"top-items","fieldId":"%s","scope":"field","count":1,"valueId":"%s","direction":"top"},
                   {"kind":"condition","family":"value","fieldId":"%s","operator":"between","value":10,"value2":50,"valueId":"%s"}],
                   "allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},
                   "values":[{"valueId":"%s","fieldId":"%s","summarizeBy":"sum"},{"valueId":"%s","fieldId":"%s","summarizeBy":"count"}],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},
                 "refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                """.formatted(region, amount, region, region, region, region, valueId, amount, valueId, valueId, amount, countValueId, amount));
        OperationMutation add = new OperationMutation("pivot.add", "sheet-1", pivot);
        JsonNode current = registry.applyPublicMutations(snapshot, List.of(add));
        JsonNode filters = current.path("sheets").get(0).path("pivots").get(0).path("layout").path("filters");
        assertEquals("field", filters.get(0).path("scope").asText());
        assertEquals("report", filters.get(1).path("scope").asText());
        assertEquals("field", filters.get(2).path("scope").asText());
        assertEquals("report", filters.get(3).path("scope").asText());
        JsonNode values = current.path("sheets").get(0).path("pivots").get(0).path("layout").path("values");
        assertEquals(2, values.size());
        assertEquals(2, values.findValues("valueId").size());

        ObjectNode duplicateValuesPivot = (ObjectNode) pivot.deepCopy();
        ArrayNode duplicateValues = (ArrayNode) duplicateValuesPivot.path("layout").path("values");
        duplicateValues.add(duplicateValues.get(0).deepCopy());
        OperationMutation duplicateValuesMutation = new OperationMutation("pivot.add", "sheet-1", duplicateValuesPivot);
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, duplicateValuesMutation, WorkbookAclRole.EDITOR));

        ObjectNode invalidScopePivot = (ObjectNode) pivot.deepCopy();
        ((ObjectNode) invalidScopePivot.path("layout").path("filters").get(0)).put("scope", "workspace");
        OperationMutation invalidScope = new OperationMutation("pivot.add", "sheet-1", invalidScopePivot);
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, invalidScope, WorkbookAclRole.EDITOR));
        assertEquals(0, snapshot.path("sheets").get(0).path("pivots").size());

        ObjectNode invalidAxisPivot = (ObjectNode) pivot.deepCopy();
        ObjectNode nonAxisFilter = (ObjectNode) invalidAxisPivot.path("layout").path("filters").get(3);
        nonAxisFilter.put("scope", "field");
        OperationMutation invalidAxis = new OperationMutation("pivot.add", "sheet-1", invalidAxisPivot);
        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, invalidAxis, WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals(0, snapshot.path("sheets").get(0).path("pivots").size());
    }

    @Test
    void pivotSourceReferencesResolveEveryCanonicalEntityAndRejectDanglingIds() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":20,"columnCount":10,"cells":{},"sheetTables":[{"id":"sheet-table-1","name":"Sales","range":{"sheetId":"sheet-1","startRow":0,"endRow":2,"startColumn":0,"endColumn":1}}],"pivots":[]},
                  {"id":"sheet-2","name":"Sheet 2","rowCount":20,"columnCount":10,"cells":{},"sheetTables":[],"pivots":[]}
                ],
                "dataModel":{"sources":[{"id":"source-1","fields":[{"id":"source-field"}]}],"tables":[],"relationships":[],"views":[]},
                "definedNameModels":[
                  {"name":"SalesData","formula":"=Sheet1!A1:B3","scope":"workbook"},
                  {"name":"SalesData","formula":"='Sheet 2'!A1:B3","scope":"sheet","sheetId":"sheet-2"}
                ]}
                """);

        for (String source : List.of(
                "{\"kind\":\"table\",\"tableId\":\"sheet-table-1\"}",
                "{\"kind\":\"named-range\",\"name\":\"SalesData\"}",
                "{\"kind\":\"named-range\",\"name\":\"SalesData\",\"sheetId\":\"sheet-2\"}",
                "{\"kind\":\"data-source\",\"dataSourceId\":\"source-1\"}")) {
            OperationMutation valid = new OperationMutation("pivot.add", "sheet-1", pivotWithSource(source, "valid-" + Math.abs(source.hashCode())));
            registry.prepare(snapshot, valid, WorkbookAclRole.EDITOR);
        }

        for (String source : List.of(
                "{\"kind\":\"table\",\"tableId\":\"missing-table\"}",
                "{\"kind\":\"named-range\",\"name\":\"MissingName\"}",
                "{\"kind\":\"named-range\",\"name\":\"SalesData\",\"sheetId\":\"missing-sheet\"}",
                "{\"kind\":\"data-source\",\"dataSourceId\":\"missing-source\"}")) {
            OperationMutation invalid = new OperationMutation("pivot.add", "sheet-1", pivotWithSource(source, "invalid-" + Math.abs(source.hashCode())));
            JsonNode before = snapshot.deepCopy();
            assertThrows(ServiceException.class, () -> registry.prepare(snapshot, invalid, WorkbookAclRole.EDITOR));
            assertEquals(before, snapshot);
        }
        ObjectNode missingSources = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) missingSources.path("dataModel")).remove("sources");
        JsonNode beforeMissingSources = missingSources.deepCopy();
        OperationMutation invalidWithoutSourceCollection = new OperationMutation("pivot.add", "sheet-1", pivotWithSource("{\"kind\":\"data-source\",\"dataSourceId\":\"missing-source\"}", "invalid-no-source-collection"));
        assertThrows(ServiceException.class, () -> registry.prepare(missingSources, invalidWithoutSourceCollection, WorkbookAclRole.EDITOR));
        assertEquals(beforeMissingSources, missingSources);
    }

    @Test
    void pivotSourceSwitchRejectsStaleFieldCatalogBeforeApplying() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":20,"columnCount":10,"cells":{},"sheetTables":[{"id":"sheet-table-1","name":"Sales","range":{"sheetId":"sheet-1","startRow":0,"endRow":2,"startColumn":0,"endColumn":0}}],"pivots":[
                  {"schema":"PivotDefinition","id":"pivot-1","source":{"kind":"table","tableId":"sheet-table-1"},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[{"fieldId":"table:sheet-table-1:column:0","name":"Amount","dataType":"number","ordinal":0}]},"layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                ]}],"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},"definedNameModels":[{"name":"SalesData","formula":"=Sheet1!A1:A3","scope":"workbook"}]}
                """);
        OperationMutation update = new OperationMutation("pivot.update", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","pivotId":"pivot-1","source":{"kind":"named-range","name":"SalesData"}}
                """));
        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, update, WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals("table", snapshot.path("sheets").get(0).path("pivots").get(0).path("source").path("kind").asText());
    }

    @Test
    void pivotNamedRangeSourceUsesExactWorkbookOrWorksheetScope() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":20,"columnCount":10,"cells":{},"pivots":[]},
                  {"id":"sheet-2","name":"Sheet 2","rowCount":20,"columnCount":10,"cells":{},"pivots":[]}
                ],
                "definedNameModels":[
                  {"name":"WorkbookOnly","formula":"=Sheet1!A1:B3","scope":"workbook"},
                  {"name":"LocalOnly","formula":"=C1:D3","scope":"sheet","sheetId":"sheet-2"}
                ]}
                """);

        OperationMutation validLocal = new OperationMutation("pivot.add", "sheet-1", pivotWithSource(
                "{\"kind\":\"named-range\",\"name\":\"LocalOnly\",\"sheetId\":\"sheet-2\"}", "pivot-local-exact"));
        registry.prepare(snapshot, validLocal, WorkbookAclRole.EDITOR);

        OperationMutation validWorkbook = new OperationMutation("pivot.add", "sheet-1", pivotWithSource(
                "{\"kind\":\"named-range\",\"name\":\"WorkbookOnly\"}", "pivot-workbook-exact"));
        registry.prepare(snapshot, validWorkbook, WorkbookAclRole.EDITOR);

        for (String source : List.of(
                "{\"kind\":\"named-range\",\"name\":\"WorkbookOnly\",\"sheetId\":\"sheet-2\"}",
                "{\"kind\":\"named-range\",\"name\":\"LocalOnly\"}")) {
            JsonNode before = snapshot.deepCopy();
            OperationMutation invalid = new OperationMutation("pivot.add", "sheet-1", pivotWithSource(source, "pivot-invalid-scope"));
            ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, invalid, WorkbookAclRole.EDITOR));
            assertEquals("NOT_FOUND", error.code());
            assertEquals(before, snapshot);
        }
    }

    private JsonNode pivotWithSource(String source, String id) throws Exception {
        return mapper.readTree("""
                {"schema":"PivotDefinition","id":"%s","source":%s,"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[]},"layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                """.formatted(id, source));
    }

    @Test
    void pivotRemovalFailsClosedWhenAChartStillReferencesThePivot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":10,"cells":{},"drawings":[
                  {"id":"pivot-chart-1","sheetId":"sheet-1","kind":"chart","payloadId":"pivot-chart-payload-1","anchor":{"kind":"absolute"},"transform":{"x":0,"y":0,"width":120,"height":80,"rotation":0},"zIndex":1}
                ],"drawingPayloads":{"pivot-chart-payload-1":{"kind":"chart","chartId":"pivot-chart-1","pivotId":"pivot-1","sourceRanges":[],"chartType":"column","elements":{"hiddenData":"show"}}},"pivots":[
                  {"schema":"PivotDefinition","id":"pivot-1","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[]},"layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                ]}]}
                """);
        OperationMutation remove = new OperationMutation("pivot.remove", "sheet-1", mapper.readTree("\"pivot-1\""));
        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, remove, WorkbookAclRole.EDITOR));
        assertEquals("CONFLICT", error.code());
    }

    @Test
    void structuralTransformsUpdateCanonicalPivotSourceAndTarget() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet 1","rowCount":10,"columnCount":10,"cells":{},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"merges":[],"hiddenRows":[],"hiddenColumns":[],"rowHeightsPx":{},"columnWidthsPx":{},"notes":[],"commentThreads":[],"drawings":[],"drawingPayloads":{},"spillRanges":[],"sheetTables":[],"conditionalFormats":[],"dataValidations":[],"protectionRules":[],"outline":{"groups":[]},"sparklines":[],"pivots":[
                  {"schema":"PivotDefinition","id":"pivot-1","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[]},"layout":{"rows":[],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                ]}]}
                """);
        OperationMutation insert = new OperationMutation("rows.inserted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","at":1,"count":1}
                """));
        JsonNode current = registry.applyPublicMutations(snapshot, List.of(insert));
        JsonNode pivot = current.path("sheets").get(0).path("pivots").get(0);
        assertEquals(2, pivot.path("source").path("range").path("endRow").asInt());
        assertEquals(5, pivot.path("target").path("anchor").path("row").asInt());
    }

    @Test
    void pivotChartCreateIsNotAWorkbookMutation() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ServiceException error = assertThrows(ServiceException.class, () -> registry.require("pivot.chart.create", false));
        assertEquals("SERVICE_UNAVAILABLE", error.code());
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
                  {"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=A2"}},"1":{"0":{"value":10}}},"pane":{"kind":"frozen","xSplit":0,"ySplit":1,"startRow":1,"startColumn":0},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"hiddenRows":[1],"rowHeightsPx":{"1":33},"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"notes":[],"commentThreads":[],"spillRanges":[],"protectionRules":[]},
                  {"id":"sheet-2","name":"Other","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=Sheet1!A2"}}},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64}
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
    void cellInsertAndRowPermutationHaveDeterministicInverseFriendlySnapshots() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=A1"}},"1":{"0":{"value":"drop"}}},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"notes":[{"row":0,"column":0,"note":{"id":"n1"}}],"commentThreads":[],"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                """);
        OperationMutation shift = new OperationMutation("cells.inserted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"operation":"insert","axis":"row","affectedBand":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":0}}
                """));
        JsonNode current = registry.prepare(snapshot, shift, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, shift);
        assertEquals("=A2", current.path("sheets").get(0).path("cells").path("1").path("0").path("formula").asText());
        assertEquals(1, current.path("sheets").get(0).path("notes").get(0).path("row").asInt());

        OperationMutation restore = new OperationMutation("cells.inserted.restore", "sheet-1", mapper.readTree("""
                {"spec":{"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"operation":"insert","axis":"row","affectedBand":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":0}},"cells":[{"row":0,"column":0,"cell":{"value":null,"formula":"=A1"}},{"row":1,"column":0,"cell":{"value":"drop"}}]}
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
