package com.xc.luckysheet.server.mutation;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.List;
import java.util.Optional;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class MutationDescriptorRegistryTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void structuralDataModelCollectionsAllowOmissionButRejectMalformedOwner() {
        ObjectNode snapshot = mapper.createObjectNode();
        assertEquals(0, SnapshotMutationSupport.dataModelArray(snapshot, "sources").size());

        snapshot.put("dataModel", "invalid");
        ServiceException error = assertThrows(ServiceException.class,
                () -> SnapshotMutationSupport.dataModelArray(snapshot, "sources"));
        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void unknownMutationsFailClosed() {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ServiceException error = assertThrows(ServiceException.class, () -> registry.require("unknown.mutation", false));
        assertEquals("VALIDATION_ERROR", error.code());
        ServiceException removedAutomation = assertThrows(ServiceException.class, () -> registry.require("automation.run", false));
        assertEquals("VALIDATION_ERROR", removedAutomation.code());
    }

    @Test
    void sheetTableRenameAppliesEveryMatchingStructuredReferenceAndRejectsMalformedReferences() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","name":"Sheet 1","rowCount":20,"columnCount":10,
                   "cells":{"0":{"0":{"value":null,"formula":"=SUM(Sales[Amount])+Sales [Other]+ÅSales[Amount]+[Book.xlsx]Sales[Amount]+IF(A1=\\"Sales[Amount]\\",0,1)"},
                     "1":{"value":null,"formulaMetadata":{"kind":"dataTable","range":"B1:B2","preservedOnly":true,"sourceFormula":"=Sales[Amount]"}}}},
                   "sheetTables":[{"id":"sales-table","sheetId":"sheet-1","name":"Sales",
                     "range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":1},
                     "hasHeaderRow":true,"hasTotalRow":false,"showBandedRows":true,"showBandedColumns":false,
                     "showFirstColumn":false,"showLastColumn":false,"showFilterButton":true,"autoExpand":"none",
                     "columns":[{"id":"amount","name":"Amount"},{"id":"other","name":"Other"}]}]},
                  {"id":"sheet-2","name":"Sheet 2","rowCount":20,"columnCount":10,
                   "cells":{"0":{"0":{"value":null,"formula":"=Sales[Amount]"}}}}
                ]}
                """);
        JsonNode original = snapshot.deepCopy();
        ObjectNode params = ((ObjectNode) snapshot.path("sheets").get(0).path("sheetTables").get(0)).deepCopy();
        params.put("name", "Orders");
        OperationMutation rename = new OperationMutation("sheetTable.update", "sheet-1", params);

        MutationApplication application = registry.require("sheetTable.update", false).applyWithPatch(snapshot, rename);

        assertEquals(3, application.structuralPatch().formulaOwnerDeltas().size());
        assertEquals("=SUM(Orders[Amount])+Orders [Other]+ÅSales[Amount]+[Book.xlsx]Sales[Amount]+IF(A1=\"Sales[Amount]\",0,1)",
                application.snapshot().path("sheets").get(0).path("cells").path("0").path("0").path("formula").asText());
        assertEquals("=Orders[Amount]", application.snapshot().path("sheets").get(0)
                .path("cells").path("0").path("1").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("B1:B2", application.snapshot().path("sheets").get(0)
                .path("cells").path("0").path("1").path("formulaMetadata").path("range").asText());
        assertEquals("=Orders[Amount]",
                application.snapshot().path("sheets").get(1).path("cells").path("0").path("0").path("formula").asText());
        assertEquals(original, snapshot);
        assertEquals(application.snapshot(), registry.applyPublicMutations(snapshot, List.of(rename)));
        JsonNode undone = registry.applyStructuralPatch(application.snapshot(), application.structuralPatch().inverse("sheetTable.update"));
        assertEquals("=Sales[Amount]", undone.path("sheets").get(0)
                .path("cells").path("0").path("1").path("formulaMetadata").path("sourceFormula").asText());
        JsonNode redone = registry.applyStructuralPatch(undone, application.structuralPatch());
        assertEquals("=Orders[Amount]", redone.path("sheets").get(0)
                .path("cells").path("0").path("1").path("formulaMetadata").path("sourceFormula").asText());

        ObjectNode malformed = snapshot.deepCopy();
        ((ObjectNode) malformed.path("sheets").get(0).path("cells").path("0").path("0")).put("formula", "=SUM(Sales[Amount)");
        JsonNode malformedOriginal = malformed.deepCopy();
        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.require("sheetTable.update", false).applyWithPatch(malformed, rename));
        assertEquals("SERVICE_UNAVAILABLE", error.code());
        assertEquals(malformedOriginal, malformed);

        ObjectNode grouped = snapshot.deepCopy();
        ObjectNode groupedCells = (ObjectNode) grouped.path("sheets").get(0).path("cells");
        ObjectNode groupedCell = groupedCells.putObject("2").putObject("0");
        groupedCell.put("formula", "=Sales[Amount]");
        groupedCell.putObject("formulaMetadata")
                .put("kind", "shared").put("sharedIndex", 7).put("sharedMaster", true)
                .put("range", "A3:A4").put("sourceFormula", "=Sales[Amount]");
        JsonNode groupedOriginal = grouped.deepCopy();
        ServiceException groupedError = assertThrows(ServiceException.class,
                () -> registry.require("sheetTable.update", false).applyWithPatch(grouped, rename));
        assertEquals("SERVICE_UNAVAILABLE", groupedError.code());
        assertEquals(groupedOriginal, grouped);

        ObjectNode missingRuleRanges = snapshot.deepCopy();
        ObjectNode missingRuleSheet = (ObjectNode) missingRuleRanges.path("sheets").get(0);
        missingRuleSheet.putArray("conditionalFormats").addObject()
                .put("id", "formula-rule-without-ranges").put("sheetId", "sheet-1")
                .put("operator", "formula").put("value1", "=Sales[Amount]");
        JsonNode missingRuleRangesOriginal = missingRuleRanges.deepCopy();
        ServiceException missingRangesError = assertThrows(ServiceException.class,
                () -> registry.require("sheetTable.update", false).applyWithPatch(missingRuleRanges, rename));
        assertEquals("VALIDATION_ERROR", missingRangesError.code());
        assertEquals(missingRuleRangesOriginal, missingRuleRanges);
    }

    @Test
    void ownedStructuralPatchReusesCandidateAndRejectsConflictsWithoutChangingBaseSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode baseSnapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1","cells":{"0":{"0":{"formula":"=A1"},"1":{"formula":"=B1"}}}}]}
                """);
        JsonNode original = baseSnapshot.deepCopy();
        StructuralPatch.CellAddress firstAddress = new StructuralPatch.CellAddress("sheet-1", 0, 0);
        StructuralPatch.CellAddress secondAddress = new StructuralPatch.CellAddress("sheet-1", 0, 1);
        StructuralPatch.FormulaOwnerDelta firstDelta = new StructuralPatch.FormulaOwnerDelta(
                "formula-cell", firstAddress, firstAddress,
                new StructuralPatch.FormulaOwnerState("=A1", null, null),
                new StructuralPatch.FormulaOwnerState("=A2", null, null));
        StructuralPatch.FormulaOwnerDelta conflictingDelta = new StructuralPatch.FormulaOwnerDelta(
                "formula-cell", secondAddress, secondAddress,
                new StructuralPatch.FormulaOwnerState("=B1", null, null),
                new StructuralPatch.FormulaOwnerState("=B2", null, null));

        ObjectNode successfulCandidate = baseSnapshot.deepCopy();
        JsonNode applied = registry.applyStructuralPatchOnOwnedSnapshot(successfulCandidate,
                new StructuralPatch(StructuralPatch.VERSION, "rows.inserted", List.of(firstDelta)));

        assertSame(successfulCandidate, applied);
        assertEquals("=A2", applied.path("sheets").get(0).path("cells").path("0").path("0").path("formula").asText());
        assertEquals(original, baseSnapshot);

        ObjectNode rejectedCandidate = baseSnapshot.deepCopy();
        ((ObjectNode) rejectedCandidate.path("sheets").get(0).path("cells").path("0").path("1")).put("formula", "=BROKEN");
        JsonNode rejectedBase = rejectedCandidate.deepCopy();
        StructuralPatch patchWithConflict = new StructuralPatch(
                StructuralPatch.VERSION, "rows.inserted", List.of(firstDelta, conflictingDelta));

        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.applyStructuralPatchOnOwnedSnapshot(rejectedCandidate, patchWithConflict));

        assertEquals("CONFLICT", error.code());
        assertTrue(error.getMessage().contains("STRUCTURAL_PATCH_PRECONDITION"));
        assertEquals("=A2", rejectedCandidate.path("sheets").get(0).path("cells").path("0").path("0").path("formula").asText());
        assertEquals("=BROKEN", rejectedCandidate.path("sheets").get(0).path("cells").path("0").path("1").path("formula").asText());
        assertEquals(original, baseSnapshot);
        assertEquals("=A1", rejectedBase.path("sheets").get(0).path("cells").path("0").path("0").path("formula").asText());
    }

    @Test
    void cellSetUsesServerResolvedRangeAndChangesSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("{\"sheets\":[{\"id\":\"sheet-1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{}}]}");
        var mutation = new OperationMutation("cell.set", "sheet-1", mapper.readTree(cellSetParams(2, 3, "{\"value\":42}", "accepted")));
        assertEquals(2, registry.resolveRanges(snapshot, mutation).get(0).startRow());
        var next = registry.applyPublicMutations(snapshot, List.of(mutation));
        assertEquals(42, next.path("sheets").get(0).path("cells").path("2").path("3").path("value").asInt());
    }

    @Test
    void publicMutationReductionPreservesInputAndEmptyBatchIsolation() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = (ObjectNode) mapper.readTree("{\"name\":\"before\",\"sheets\":[{\"id\":\"sheet-1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{}}]}");
        JsonNode original = snapshot.deepCopy();
        var mutation = new OperationMutation("cell.set", "sheet-1", mapper.readTree(cellSetParams(2, 3, "{\"value\":42}", "accepted")));

        JsonNode next = registry.applyPublicMutations(snapshot, List.of(mutation));

        assertNotSame(snapshot, next);
        assertEquals(original, snapshot);
        assertEquals(42, next.path("sheets").get(0).path("cells").path("2").path("3").path("value").asInt());

        JsonNode emptyReduction = registry.applyPublicMutations(snapshot, List.of());
        assertNotSame(snapshot, emptyReduction);
        assertEquals(snapshot, emptyReduction);
        ((ObjectNode) emptyReduction).put("isolationProbe", true);
        assertTrue(snapshot.get("isolationProbe") == null);
    }

    @Test
    void definedNameMutationsKeepCaseInsensitiveProjectionSynchronized() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1"}],
                 "definedNames":{"TaxRate":"0.1"},
                 "definedNameModels":[{"name":"TaxRate","formula":"0.1","scope":"workbook"}]}
                """);

        JsonNode changed = registry.applyPublicMutations(snapshot, List.of(new OperationMutation(
                "name.set", "sheet-1", mapper.readTree("""
                        {"model":{"name":"taxrate","formula":"0.2","scope":"workbook"}}
                        """))));

        assertEquals(1, changed.path("definedNameModels").size());
        assertEquals("taxrate", changed.path("definedNameModels").get(0).path("name").asText());
        assertEquals(1, changed.path("definedNames").size());
        assertEquals("0.2", changed.path("definedNames").path("taxrate").asText());

        JsonNode removed = registry.applyPublicMutations(changed, List.of(new OperationMutation(
                "name.remove", "sheet-1", mapper.readTree("""
                        {"name":"TAXRATE","scope":"workbook"}
                        """))));

        assertEquals(0, removed.path("definedNameModels").size());
        assertEquals(0, removed.path("definedNames").size());
    }

    @Test
    void structuralDescriptorKeepsPublicPurityAndMutatesOnlyOwnedSnapshots() throws Exception {
        StructuralMutationDescriptor descriptor = new StructuralMutationDescriptor("rows.inserted");
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,
                  "cells":{"0":{"0":{"value":null,"formula":"=A1"}}},"pane":{"kind":"none"},
                  "defaultRowHeightPx":20,"defaultColumnWidthPx":64,"hiddenRows":[],"hiddenColumns":[],
                  "rowHeightsPx":{},"columnWidthsPx":{},"merges":[],"conditionalFormats":[],"dataValidations":[],
                  "pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],
                  "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                  "spillRanges":[],"protectionRules":[],"outline":{"groups":[]}}],
                 "definedNames":{"TaxRate":"=A2"},
                 "definedNameModels":[{"name":"TaxRate","formula":"=A2","scope":"workbook",
                   "anchor":{"sheetId":"sheet-1","row":3,"column":0}}]}
                """);
        JsonNode original = snapshot.deepCopy();
        OperationMutation mutation = new OperationMutation("rows.inserted", "sheet-1",
                mapper.readTree("{\"sheetId\":\"sheet-1\",\"at\":0,\"count\":1}"));

        MutationApplication standalone = descriptor.applyWithPatch(snapshot, mutation);
        assertNotSame(snapshot, standalone.snapshot());
        assertEquals(original, snapshot);
        assertEquals("=A2", standalone.snapshot().path("sheets").get(0)
                .path("cells").path("1").path("0").path("formula").asText());
        assertEquals("=A3", standalone.snapshot().path("definedNameModels").get(0).path("formula").asText());
        assertEquals(4, standalone.snapshot().path("definedNameModels").get(0).path("anchor").path("row").asInt());
        assertEquals("=A2", snapshot.path("definedNameModels").get(0).path("formula").asText());

        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode reduced = registry.applyPublicMutations(snapshot, List.of(mutation, mutation));
        assertNotSame(snapshot, reduced);
        assertEquals(original, snapshot);
        assertEquals("=A3", reduced.path("sheets").get(0)
                .path("cells").path("2").path("0").path("formula").asText());
        assertEquals("=A4", reduced.path("definedNameModels").get(0).path("formula").asText());
        assertEquals(5, reduced.path("definedNameModels").get(0).path("anchor").path("row").asInt());

        ObjectNode ownedSnapshot = snapshot.deepCopy();
        MutationApplication owned = descriptor.applyWithPatchOnOwnedSnapshot(ownedSnapshot, mutation);
        assertSame(ownedSnapshot, owned.snapshot());
        assertEquals("=A2", ownedSnapshot.path("sheets").get(0)
                .path("cells").path("1").path("0").path("formula").asText());
        assertEquals("=A3", ownedSnapshot.path("definedNameModels").get(0).path("formula").asText());
        assertEquals(4, ownedSnapshot.path("definedNameModels").get(0).path("anchor").path("row").asInt());
        assertEquals(original, snapshot);
    }

    @Test
    void structuralPatchMigrationPreservesPatchlessMutationSlots() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("{\"sheets\":[{\"id\":\"sheet-1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{}}]}");
        List<OperationMutation> mutations = List.of(
                new OperationMutation("cell.set", "sheet-1", mapper.readTree(cellSetParams(1, 1, "{\"value\":1}", "accepted"))),
                new OperationMutation("cell.set", "sheet-1", mapper.readTree(cellSetParams(2, 2, "{\"value\":2}", "accepted"))));

        var replay = registry.replayStructuralPatchesForMigration(snapshot, mutations, Collections.nCopies(mutations.size(), null));

        assertEquals(List.of(Optional.empty(), Optional.empty()), replay.structuralPatches());
        assertEquals(1, replay.snapshot().path("sheets").get(0).path("cells").path("1").path("1").path("value").asInt());
        assertEquals(2, replay.snapshot().path("sheets").get(0).path("cells").path("2").path("2").path("value").asInt());
    }

    @Test
    void structuralPatchDefinedNameParsingRejectsInvalidAnchorCoordinates() throws Exception {
        JsonNode fractionalRow = mapper.readTree("""
                [{"name":"LocalName","formula":"=A1","scope":"workbook",
                  "anchor":{"sheetId":"sheet-1","row":1.5,"column":0}}]
                """);
        JsonNode outOfBoundsColumn = mapper.readTree("""
                [{"name":"LocalName","formula":"=A1","scope":"workbook",
                  "anchor":{"sheetId":"sheet-1","row":0,"column":16384}}]
                """);

        ServiceException fractionalError = assertThrows(ServiceException.class,
                () -> StructuralSnapshotReducer.definedNameOwnerDeltas(fractionalRow, fractionalRow));
        ServiceException boundsError = assertThrows(ServiceException.class,
                () -> StructuralSnapshotReducer.definedNameOwnerDeltas(outOfBoundsColumn, outOfBoundsColumn));

        assertEquals("VALIDATION_ERROR", fractionalError.code());
        assertEquals("VALIDATION_ERROR", boundsError.code());
    }

    @Test
    void structuralMutationRejectsProjectionOnlyDefinedNamesWithoutChangingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,
                  "cells":{},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                  "hiddenRows":[],"hiddenColumns":[],"rowHeightsPx":{},"columnWidthsPx":{},"merges":[],
                  "conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],
                  "drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[],"outline":{"groups":[]}}],
                 "definedNames":{"TaxRate":"=A2"}}
                """);
        JsonNode original = snapshot.deepCopy();
        OperationMutation mutation = new OperationMutation("rows.inserted", "sheet-1",
                mapper.readTree("{\"sheetId\":\"sheet-1\",\"at\":0,\"count\":1}"));

        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(snapshot, List.of(mutation)));

        assertEquals("SERVICE_UNAVAILABLE", error.code());
        assertEquals(original, snapshot);
    }

    @Test
    void committedStructuralReplayUsesDetachedBatchAndRejectsCorruptPatchWithoutChangingInput() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,
                  "cells":{"0":{"0":{"value":null,"formula":"=A1"}}},"pane":{"kind":"none"},
                  "defaultRowHeightPx":20,"defaultColumnWidthPx":64,"hiddenRows":[],"hiddenColumns":[],
                  "rowHeightsPx":{},"columnWidthsPx":{},"merges":[],"conditionalFormats":[],"dataValidations":[],
                  "pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],
                  "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                  "spillRanges":[],"protectionRules":[],"outline":{"groups":[]}}],
                 "definedNames":{"TaxRate":"=A2"},
                 "definedNameModels":[{"name":"TaxRate","formula":"=A2","scope":"workbook",
                   "anchor":{"sheetId":"sheet-1","row":3,"column":0}}]}
                """);
        JsonNode original = snapshot.deepCopy();
        OperationMutation mutation = new OperationMutation("rows.inserted", "sheet-1",
                mapper.readTree("{\"sheetId\":\"sheet-1\",\"at\":0,\"count\":1}"));
        MutationApplication generated = new StructuralMutationDescriptor("rows.inserted")
                .applyWithPatch(snapshot, mutation);
        StructuralPatch patch = generated.structuralPatch();
        List<RangeRef> affectedRanges = registry.resolveRanges(snapshot, mutation);
        List<RangeRef> impactRanges = registry.structuralImpactRanges(patch);
        CommittedOperationMutation committed = CommittedOperationMutation.from(
                mutation, affectedRanges, impactRanges, patch);

        JsonNode replayed = registry.applyCommittedMutations(snapshot, List.of(committed),
                Collections.singletonList(null));
        assertNotSame(snapshot, replayed);
        assertEquals(generated.snapshot(), replayed);
        assertEquals(1, patch.definedNameOwnerDeltas().size());
        StructuralPatch.DefinedNameOwnerDelta nameDelta = patch.definedNameOwnerDeltas().get(0);
        assertEquals("=A2", nameDelta.before().formula());
        assertEquals("=A3", nameDelta.after().formula());
        assertEquals(3, nameDelta.before().anchor().row());
        assertEquals(4, nameDelta.after().anchor().row());
        assertEquals("=A3", replayed.path("definedNames").path("TaxRate").asText());

        StructuralPatch nameOnlyInverse = new StructuralPatch(StructuralPatch.VERSION, "rows.deleted", List.of(),
                List.of(nameDelta.inverse()));
        JsonNode restoredName = registry.applyStructuralPatch(replayed, nameOnlyInverse);
        assertEquals("=A2", restoredName.path("definedNameModels").get(0).path("formula").asText());
        assertEquals(3, restoredName.path("definedNameModels").get(0).path("anchor").path("row").asInt());
        assertEquals("=A2", restoredName.path("definedNames").path("TaxRate").asText());

        assertEquals(original, snapshot);

        StructuralPatch corruptPatch = new StructuralPatch(StructuralPatch.VERSION, "rows.deleted",
                patch.formulaOwnerDeltas());
        CommittedOperationMutation corrupt = CommittedOperationMutation.from(
                mutation, affectedRanges, registry.structuralImpactRanges(corruptPatch), corruptPatch);
        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.applyCommittedMutations(snapshot, List.of(corrupt), Collections.singletonList(null)));
        assertEquals("STORAGE_CORRUPT", error.code());
        assertEquals(original, snapshot);
    }

    @Test
    void cellSetExtendsTheCanonicalWorksheetExtent() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("{\"sheets\":[{\"id\":\"sheet-1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{}}]}");
        var mutation = new OperationMutation("cell.set", "sheet-1", mapper.readTree(cellSetParams(1_000, 26, "{\"value\":42}", "accepted")));

        var next = registry.applyPublicMutations(snapshot, List.of(mutation));

        assertEquals(1_001, next.path("sheets").get(0).path("rowCount").asInt());
        assertEquals(27, next.path("sheets").get(0).path("columnCount").asInt());
    }

    @Test
    void sheetExtentGrowCommitsMonotonicallyAndRejectsShrinkOrOverflow() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("{\"sheets\":[{\"id\":\"sheet-1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{}}]}");
        var grow = new OperationMutation("sheet.extent.grow", "sheet-1", mapper.readTree("{\"sheetId\":\"sheet-1\",\"rowCount\":2000,\"columnCount\":52}"));
        var next = registry.applyPublicMutations(snapshot, List.of(grow));
        assertEquals(2000, next.path("sheets").get(0).path("rowCount").asInt());
        assertEquals(52, next.path("sheets").get(0).path("columnCount").asInt());

        var shrink = new OperationMutation("sheet.extent.grow", "sheet-1", mapper.readTree("{\"sheetId\":\"sheet-1\",\"rowCount\":999,\"columnCount\":52}"));
        var overflow = new OperationMutation("sheet.extent.grow", "sheet-1", mapper.readTree("{\"sheetId\":\"sheet-1\",\"rowCount\":1048577,\"columnCount\":52}"));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(shrink)));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(overflow)));
    }

    @Test
    void workbookEditingOptionsUseOneValidatedRootReducer() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("{\"editingOptions\":{\"allowEditDirectly\":true,\"moveAfterEnter\":true,\"enterDirection\":\"down\",\"formulaAutoComplete\":true,\"valueAutoComplete\":true,\"fixedDecimalPlaces\":null},\"sheets\":[{\"id\":\"sheet-1\",\"cells\":{}}]}");
        var options = mapper.readTree("{\"allowEditDirectly\":false,\"moveAfterEnter\":true,\"enterDirection\":\"right\",\"formulaAutoComplete\":true,\"valueAutoComplete\":false,\"fixedDecimalPlaces\":2}");
        var mutation = new OperationMutation("workbook.editing.options.set", "sheet-1", options);
        assertEquals(List.of(), registry.resolveRanges(snapshot, mutation));
        var next = registry.applyPublicMutations(snapshot, List.of(mutation));
        assertEquals("right", next.path("editingOptions").path("enterDirection").asText());
        assertEquals(2, next.path("editingOptions").path("fixedDecimalPlaces").asInt());
        var invalid = new OperationMutation("workbook.editing.options.set", "sheet-1", mapper.readTree("{\"allowEditDirectly\":true,\"moveAfterEnter\":true,\"enterDirection\":\"down\",\"formulaAutoComplete\":true,\"fixedDecimalPlaces\":null}"));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(invalid)));
    }

    @Test
    void cellSetRevalidatesTheBoundCandidateAgainstTheCommitSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{},"dataValidations":[{
                  "id":"whole-positive","type":"whole","operator":"greaterThan","formula1":"0","alertStyle":"stop","allowBlank":false,
                  "ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0}]
                }]}]}
                """);
        var invalid = new OperationMutation("cell.set", "sheet-1", mapper.readTree(cellSetParams(0, 0, "{\"value\":-1}", "accepted")));
        ServiceException validation = assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(invalid)));
        assertEquals("VALIDATION_ERROR", validation.code());

        var mismatched = new OperationMutation("cell.set", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","row":0,"column":0,"value":{"value":2},"writeAuthority":{"kind":"script",
                 "target":{"sheetId":"sheet-1","row":0,"column":0},"candidate":{"value":3},"validationDecision":{"status":"accepted","ruleId":"whole-positive","alertStyle":"stop"}}}
                """));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(mismatched)));
    }

    @Test
    void hyperlinkSetAcceptsTypedTargetsAndRejectsUnknownDestinations() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"definedNameModels":[{"name":"SalesTotal","formula":"=Sheet1!A1","scope":"workbook"}],
                 "sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":8,"columnCount":8,"cells":{},"hyperlinks":[]}]}
                """);
        var valid = new OperationMutation("hyperlink.set", "sheet-1", mapper.readTree("""
                {"row":0,"column":0,"hyperlink":{"id":"link-1","target":{"kind":"sheet","sheetId":"sheet-1","address":"B2"},"tooltip":"Open"}}
                """));
        var next = registry.applyPublicMutations(snapshot, List.of(valid));
        assertEquals("sheet", next.path("sheets").get(0).path("hyperlinks").get(0).path("hyperlink").path("target").path("kind").asText());
        var invalid = new OperationMutation("hyperlink.set", "sheet-1", mapper.readTree("""
                {"row":0,"column":0,"hyperlink":{"id":"link-2","target":{"kind":"sheet","sheetId":"missing","address":"A1"}}}
                """));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(next, List.of(invalid)));
    }

    @Test
    void rangePasteAppliesCanonicalSnapshotAndClearsABoundedSource() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":20,"cells":{"0":{"0":{"value":"move"}}}}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":1,"column":1},"sourceExtent":{"rows":1,"columns":1},"clipboard":{"schema":"SparseClipboardPayload","transfer":"move","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"sourceExtent":{"rows":1,"columns":1},"occupiedCells":[],"rangeMetadata":{"columnWidths":[],"validations":[],"conditionalFormats":[],"notes":[],"comments":[],"hyperlinks":[]}},
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
    void crossSheetRangePasteFailsClosedBeforeApplyingSnapshots() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","rowCount":20,"columnCount":20,"cells":{"0":{"0":{"value":"source"}}}},
                  {"id":"sheet-2","rowCount":20,"columnCount":20,"cells":{"1":{"1":{"value":"target"}}}}
                ]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-2", mapper.readTree("""
                {"sheetId":"sheet-2","targetOrigin":{"row":2,"column":2},"sourceExtent":{"rows":1,"columns":1},
                 "clipboard":{"schema":"SparseClipboardPayload","transfer":"move","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"sourceExtent":{"rows":1,"columns":1},"occupiedCells":[],"rangeMetadata":{"columnWidths":[],"validations":[],"conditionalFormats":[],"notes":[],"comments":[],"hyperlinks":[]}},
                 "transfer":"move","clearSource":true,"sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},
                 "sourceSnapshot":{"cells":[]},
                 "spec":{"content":"all","formatting":"all","metadata":{"commentsNotes":true,"validation":true,"columnWidths":false,"conditionalFormats":true,"hyperlinks":true},"operation":"none","skipBlanks":false,"transpose":false,"link":false},
                 "snapshot":{"cells":[{"row":2,"column":2,"value":{"value":"source"}}]}}
                """));

        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));

        assertEquals("SERVICE_UNAVAILABLE", error.code());
        assertTrue(error.getMessage().contains("UNSUPPORTED_FEATURE"));
        assertEquals("source", snapshot.path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());
        assertEquals("target", snapshot.path("sheets").get(1).path("cells").path("1").path("1").path("value").asText());
    }

    @Test
    void rangePasteRejectsMoveSourceRangeDifferentFromClipboardRange() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":20,"cells":{"0":{"0":{"value":"source"}}}}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":2,"column":2},"sourceExtent":{"rows":1,"columns":1},
                 "clipboard":{"schema":"SparseClipboardPayload","transfer":"move","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"sourceExtent":{"rows":1,"columns":1},"occupiedCells":[],"rangeMetadata":{"columnWidths":[],"validations":[],"conditionalFormats":[],"notes":[],"comments":[],"hyperlinks":[]}},
                 "transfer":"move","clearSource":true,"sourceRange":{"sheetId":"sheet-1","startRow":1,"endRow":1,"startColumn":0,"endColumn":0},
                 "spec":{"content":"all","formatting":"all","metadata":{"commentsNotes":true,"validation":true,"columnWidths":false,"conditionalFormats":true,"hyperlinks":true},"operation":"none","skipBlanks":false,"transpose":false,"link":false},
                 "snapshot":{"clearRanges":[{"sheetId":"sheet-1","startRow":2,"endRow":2,"startColumn":2,"endColumn":2}],"cells":[]}}
                """));

        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));

        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals("source", snapshot.path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());
    }

    private static String cellSetParams(int row, int column, String value, String status) {
        return "{\"sheetId\":\"sheet-1\",\"row\":" + row + ",\"column\":" + column + ",\"value\":" + value
                + ",\"writeAuthority\":{\"kind\":\"script\",\"target\":{\"sheetId\":\"sheet-1\",\"row\":" + row
                + ",\"column\":" + column + "},\"candidate\":" + value + ",\"validationDecision\":{\"status\":\"" + status + "\"}}}";
    }

    @Test
    void rangePasteRejectsAnOversizedSourceBeforeApplyingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":20,"cells":{"0":{"0":{"value":"keep"}}}}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":1,"column":1},"sourceExtent":{"rows":1,"columns":1},"clipboard":{"schema":"SparseClipboardPayload","transfer":"move","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"sourceExtent":{"rows":1,"columns":1},"occupiedCells":[],"rangeMetadata":{"columnWidths":[],"validations":[],"conditionalFormats":[],"notes":[],"comments":[],"hyperlinks":[]}},
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
    void clearFormatsUsesFamilyAndCropsConditionalFormatRangesExactly() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{"2":{"2":{"value":"keep","style":{"bold":true}}}},
                  "conditionalFormats":[{"id":"cf-1","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":4}],"type":"highlight"}]}]}
                """);
        var mutation = new OperationMutation("range.clear", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":1,"endRow":3,"startColumn":1,"endColumn":3},"family":"formats"}
                """));
        var next = registry.applyPublicMutations(snapshot, List.of(mutation));
        assertEquals(4, next.path("sheets").get(0).path("conditionalFormats").get(0).path("ranges").size());
        assertEquals("keep", next.path("sheets").get(0).path("cells").path("2").path("2").path("value").asText());
        assertEquals(true, next.path("sheets").get(0).path("cells").path("2").path("2").path("style").isMissingNode());

        var legacy = new OperationMutation("range.clear", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"mode":"formats"}
                """));
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, legacy, WorkbookAclRole.EDITOR));
    }

    @Test
    void clearContentsRemovesFormulaDefinitionProvenanceAndCachedResult() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":1,"columnCount":1,"cells":{"0":{"0":{
                  "value":7,"formula":"=1+6","formulaValue":7,"displayValue":"7",
                  "formulaMetadata":{"kind":"normal","sourceFormula":"=1+6"}
                }}}}]}
                """);
        var mutation = new OperationMutation("range.clear", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"family":"contents"}
                """));

        var next = registry.applyPublicMutations(snapshot, List.of(mutation));
        var cell = next.path("sheets").get(0).path("cells").path("0").path("0");
        assertTrue(cell.path("value").isNull());
        assertTrue(cell.path("formula").isMissingNode());
        assertTrue(cell.path("formulaValue").isMissingNode());
        assertTrue(cell.path("formulaMetadata").isMissingNode());
        assertTrue(cell.path("displayValue").isMissingNode());
    }

    @Test
    void metadataOnlyClearRestorePreservesCellStorageWhenCellSnapshotIsOmitted() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{"0":{"0":{"value":"keep"}}},
                  "hyperlinks":[],"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}]}
                """);
        var mutation = new OperationMutation("range.clear.restore", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","family":"hyperlinks","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},
                 "snapshot":{"notes":[],"comments":[],"hyperlinks":[{"row":0,"column":0,"hyperlink":{"id":"link","target":{"kind":"url","url":"https://example.com"}}}]}}
                """));

        var next = registry.applyPublicMutations(snapshot, List.of(mutation));

        assertEquals("keep", next.path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());
        assertEquals("link", next.path("sheets").get(0).path("hyperlinks").get(0).path("hyperlink").path("id").asText());
    }

    @Test
    void clearRestoreRejectsMissingCellSnapshotForCellClearingFamily() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{"0":{"0":{"value":"keep"}}},
                  "hyperlinks":[],"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}]}
                """);
        var mutation = new OperationMutation("range.clear.restore", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","family":"contents","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},
                 "snapshot":{"notes":[],"comments":[],"hyperlinks":[]}}
                """));

        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(snapshot, List.of(mutation)));

        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void clearRestoreRejectsMissingMetadataSnapshotBeforeRemovingExistingReviewData() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{},"hyperlinks":[],
                  "review":{"notesByCell":{"0:0":"keep"},"notesById":{"note-1":{"id":"note-1","sheetId":"sheet-1","row":0,"column":0,"content":"keep"}},
                  "threadIdsByCell":{},"threadsById":{}}}]}
                """);
        var mutation = new OperationMutation("range.clear.restore", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","family":"hyperlinks","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},
                 "snapshot":{"hyperlinks":[],"comments":[]}}
                """));

        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(snapshot, List.of(mutation)));

        assertEquals("VALIDATION_ERROR", error.code());
    }

    @Test
    void rangePasteAcceptsMetadataWhoseEveryOwnedRangeIsInsideTheTarget() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{},"dataValidations":[]}]}
                """);
        var mutation = new OperationMutation("range.paste", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","targetOrigin":{"row":0,"column":0},"sourceExtent":{"rows":2,"columns":1},"clipboard":{"schema":"SparseClipboardPayload","transfer":"copy","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceExtent":{"rows":2,"columns":1},"occupiedCells":[],"rangeMetadata":{"columnWidths":[],"validations":[],"conditionalFormats":[],"notes":[],"comments":[],"hyperlinks":[]}},
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
                {"sheetId":"sheet-1","targetOrigin":{"row":0,"column":0},"sourceExtent":{"rows":1,"columns":1},"clipboard":{"schema":"SparseClipboardPayload","transfer":"copy","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"sourceExtent":{"rows":1,"columns":1},"occupiedCells":[],"rangeMetadata":{"columnWidths":[],"validations":[],"conditionalFormats":[],"notes":[],"comments":[],"hyperlinks":[]}},
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
                {"sheetId":"sheet-1","targetOrigin":{"row":0,"column":0},"sourceExtent":{"rows":1,"columns":1},"clipboard":{"schema":"SparseClipboardPayload","transfer":"copy","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"sourceExtent":{"rows":1,"columns":1},"occupiedCells":[],"rangeMetadata":{"columnWidths":[],"validations":[],"conditionalFormats":[],"notes":[],"comments":[],"hyperlinks":[]}},
                 "transfer":"copy","clearSource":false,
                 "spec":{"content":"all","formatting":"all","metadata":{"commentsNotes":true,"validation":true,"columnWidths":true,"conditionalFormats":true,"hyperlinks":true},"operation":"none","skipBlanks":false,"transpose":false,"link":false},
                 "snapshot":{"cells":[],"columnWidths":[{"column":0,"widthPx":120}]}}
                """));

        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));
        assertEquals("FORBIDDEN", error.code());

        var owner = registry.prepare(snapshot, mutation, WorkbookAclRole.OWNER);
        assertEquals(9, owner.affectedRanges().get(1).endRow());
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
                "style.set", "style.preset.set", "merge.set", "merge.remove", "freeze.set", "row.resize", "column.resize", "column.defaultWidth.resize", "columns.visibility", "view.set", "sheet.hidden", "sheet.unhidden", "sheet.tabColor",
                "note.set", "note.remove", "note.visibility", "comment.add", "comment.reply", "comment.reply.remove", "comment.resolve", "comment.remove",
                "sheet.protect.set", "sheet.protect.remove", "sheet.extent.grow", "workbook.renamed", "workbook.editing.options.set",
                "sheet.add", "sheet.remove", "sheet.rename", "sheet.duplicated", "sheet.restore", "hyperlink.set", "hyperlink.remove",
                "sheet.reordered",
                "row.hidden", "row.unhidden", "rows.unhidden.all", "rows.hidden.restore",
                "column.hidden", "column.unhidden", "columns.unhidden.all", "columns.hidden.restore",
                "autoFilter.set", "autoFilter.remove", "cf.add", "cf.remove", "cf.clear", "cf.reorder", "dv.add", "dv.remove", "banded.set", "outline.set",
                "sheetTable.add", "sheetTable.remove", "sheetTable.update", "sheetTable.autoFilter.set", "tableSheet.update", "ganttSheet.update", "reportSheet.update",
                "drawing.add", "drawing.remove", "drawing.transform", "drawing.transform.batch", "drawing.anchor", "drawing.payload.update", "drawing.zorder", "drawing.zorder.restore", "drawing.visibility.set", "drawing.rename",
                "pivot.add", "pivot.remove", "pivot.update", "pivot.refresh", "pivot.drilldown.add", "pivot.drilldown.remove",
                "sparkline.add", "sparkline.remove", "sparkline.update", "sparkline.group.add", "sparkline.group.remove", "sparkline.group.replace",
                "table.add", "table.remove", "name.set", "name.remove", "workbook.calculation.mode.set",
                "pageLayout.margins.set", "pageLayout.orientation.set", "pageLayout.paperSize.set", "pageLayout.pageSetupDetail.set", "pageLayout.scaleToFit.set", "pageLayout.printTitles.set", "pageLayout.printArea.set", "pageLayout.printArea.clear", "pageLayout.pageBreak.insert", "pageLayout.pageBreak.remove", "pageLayout.pageBreak.clear", "pageLayout.printGridlines.set", "pageLayout.printHeadings.set", "pageLayout.viewGridlines.set", "pageLayout.viewHeadings.set"
                , "query.definition.replace", "query.load.range", "query.load.sheet-table", "query.load.pivot-source", "query.load.workbook-table",
                "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted", "cells.inserted", "cells.deleted", "cells.inserted.restore", "cells.deleted.restore", "rows.permuted", "range.move", "rows.visibility",
                "fill.applied", "fill.restored",
                "dataSource.add", "dataSource.update", "dataSource.remove", "dataRegion.add", "dataRegion.remove", "analysis.view.replace"
        ), Set.copyOf(registry.acceptedIds()));
    }

    @Test
    void analysisViewReplacePersistsSharedDashboardStateAndRemovesItAtomically() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[{"id":"table-1","fields":[{"id":"region"},{"id":"amount"}]}],"relationships":[],"views":[]},
                 "sheets":[{"id":"sheet-1","rowCount":20,"columnCount":5,"cells":{}}]}
                """);
        var view = mapper.readTree("""
                {"view":{"kind":"analysis","id":"analysis-1","name":"Sales dashboard","tableId":"table-1",
                  "fields":[{"fieldId":"region","caption":"Region"},{"fieldId":"amount","caption":"Amount"}],
                  "filters":[{"id":"filter-region","fieldId":"region","operator":"in","values":["East"]}],
                  "charts":[{"chartId":"chart-1","fieldMap":{"category":"region","value":"amount"}}],
                  "layout":{"columns":2,"rowHeightPx":240,"gapPx":12},"revision":1}}
                """);
        var add = new OperationMutation("analysis.view.replace", "sheet-1", view);
        var added = registry.applyPublicMutations(snapshot, List.of(add));
        assertEquals("analysis-1", added.path("dataModel").path("views").get(0).path("id").asText());
        assertEquals("in", added.path("dataModel").path("views").get(0).path("filters").get(0).path("operator").asText());

        var remove = new OperationMutation("analysis.view.replace", "sheet-1", mapper.readTree("{\"view\":null,\"viewId\":\"analysis-1\"}"));
        var removed = registry.applyPublicMutations(added, List.of(remove));
        assertEquals(0, removed.path("dataModel").path("views").size());
    }

    @Test
    void analysisViewRejectsUnknownTableFieldWithoutChangingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[{"id":"table-1","fields":[{"id":"region"}]}],"relationships":[],"views":[]},
                 "sheets":[{"id":"sheet-1","rowCount":20,"columnCount":5,"cells":{}}]}
                """);
        var invalid = new OperationMutation("analysis.view.replace", "sheet-1", mapper.readTree("""
                {"view":{"kind":"analysis","id":"analysis-1","name":"Broken","tableId":"table-1","fields":[{"fieldId":"missing","caption":"Missing"}],
                  "filters":[],"charts":[],"layout":{"columns":1,"rowHeightPx":200,"gapPx":8},"revision":0}}
                """));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(invalid)));
        assertEquals(0, snapshot.path("dataModel").path("views").size());
    }

    @Test
    void analysisViewRejectsIncompleteChartFieldMapWithoutChangingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[{"id":"table-1","fields":[{"id":"region"},{"id":"amount"}]}],"relationships":[],"views":[]},
                 "sheets":[{"id":"sheet-1","rowCount":20,"columnCount":5,"cells":{}}]}
                """);
        var invalid = new OperationMutation("analysis.view.replace", "sheet-1", mapper.readTree("""
                {"view":{"kind":"analysis","id":"analysis-1","name":"Broken chart","tableId":"table-1","fields":[{"fieldId":"region","caption":"Region"},{"fieldId":"amount","caption":"Amount"}],
                  "filters":[],"charts":[{"chartId":"chart-1","fieldMap":{"category":"region"}}],"layout":{"columns":1,"rowHeightPx":200,"gapPx":8},"revision":0}}
                """));
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(invalid)));
        assertEquals(0, snapshot.path("dataModel").path("views").size());
    }

    @Test
    void analysisViewExpectedRevisionRejectsStaleDashboardEdit() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[{"id":"table-1","fields":[{"id":"region"}]}],"relationships":[],"views":[]},
                 "sheets":[{"id":"sheet-1","rowCount":20,"columnCount":5,"cells":{}}]}
                """);
        var create = new OperationMutation("analysis.view.replace", "sheet-1", mapper.readTree("""
                {"view":{"kind":"analysis","id":"analysis-1","name":"Dashboard","tableId":"table-1","fields":[{"fieldId":"region","caption":"Region"}],"filters":[],"charts":[],"layout":{"columns":1,"rowHeightPx":200,"gapPx":8},"revision":0},"expectedRevision":null}
                """));
        var current = registry.applyPublicMutations(snapshot, List.of(create));
        var update = new OperationMutation("analysis.view.replace", "sheet-1", mapper.readTree("""
                {"view":{"kind":"analysis","id":"analysis-1","name":"Dashboard v2","tableId":"table-1","fields":[{"fieldId":"region","caption":"Region"}],"filters":[],"charts":[],"layout":{"columns":1,"rowHeightPx":200,"gapPx":8},"revision":1},"expectedRevision":0}
                """));
        current = registry.applyPublicMutations(current, List.of(update));
        var stale = new OperationMutation("analysis.view.replace", "sheet-1", mapper.readTree("""
                {"view":{"kind":"analysis","id":"analysis-1","name":"Stale","tableId":"table-1","fields":[{"fieldId":"region","caption":"Region"}],"filters":[],"charts":[],"layout":{"columns":1,"rowHeightPx":200,"gapPx":8},"revision":1},"expectedRevision":0}
                """));
        var before = current.deepCopy();
        ServiceException conflict = assertThrows(ServiceException.class, () -> registry.applyPublicMutations(before, List.of(stale)));
        assertEquals("CONFLICT", conflict.code());
        assertEquals(before, current);
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
                "pivot.chart.create",
                "workbook.restore"
        ), registry.unavailableReasons().keySet());
    }

    @Test
    void commenterMayCommitReviewMutationButCannotWriteCells() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","cells":{},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}]}
                """);
        var note = new OperationMutation("note.set", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","row":2,"column":3,"note":{"id":"n-1","author":"guest","text":"Review","createdAt":"2026-08-23T00:00:00Z","visible":true}}
                """));

        var prepared = registry.prepare(snapshot, note, WorkbookAclRole.COMMENTER);
        assertEquals(2, prepared.affectedRanges().get(0).startRow());
        var next = prepared.descriptor().apply(snapshot, note);
        assertEquals("n-1", next.path("sheets").get(0).path("review").path("notesById").path("n-1").path("id").asText());

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
                {"sheets":[{"id":"sheet-1","rowCount":10,"columnCount":10,"cells":{},"protectionRules":[
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
    void sheetProtectionUsesCellLockedStyleAndNativeAllowFlags() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":4,"columnCount":4,
                  "cells":{"0":{"0":{"value":"unlocked","style":{"locked":false}},"1":{"value":"locked","style":{"locked":true}}}},
                  "protectionRules":[{"id":"sheet-lock","scope":"sheet","sheetId":"sheet-1","locked":true,"allow":{"formatCells":true}}]}]}
                """);

        var unlocked = new OperationMutation("cell.set", "sheet-1", mapper.readTree("""
                {"row":0,"column":0,"value":{"value":"changed"}}
                """));
        assertDoesNotThrow(() -> registry.prepare(snapshot, unlocked, WorkbookAclRole.EDITOR));

        var locked = new OperationMutation("cell.set", "sheet-1", mapper.readTree("""
                {"row":0,"column":1,"value":{"value":"rejected"}}
                """));
        ServiceException lockedError = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, locked, WorkbookAclRole.EDITOR));
        assertEquals("FORBIDDEN", lockedError.code());

        var format = new OperationMutation("style.set", "sheet-1", mapper.readTree("""
                {"range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":1,"endColumn":1},"style":{"bold":true}}
                """));
        assertDoesNotThrow(() -> registry.prepare(snapshot, format, WorkbookAclRole.EDITOR));
    }

    @Test
    void mixedLockedCellsFailBeforeACollectionMutationCanBeApplied() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        var snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":4,"columnCount":4,
                  "cells":{"0":{"0":{"value":"open","style":{"locked":false}},"1":{"value":"closed","style":{"locked":true}}}},
                  "protectionRules":[{"id":"sheet-lock","scope":"sheet","sheetId":"sheet-1","locked":true,"allow":{}}]}]}
                """);
        var mutation = new OperationMutation("range.set", "sheet-1", mapper.readTree("""
                {"startRow":0,"startColumn":0,"values":[[{"value":"a"},{"value":"b"}]]}
                """));
        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));
        assertEquals("FORBIDDEN", error.code());
        assertEquals("open", snapshot.path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());
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
                "{\"kind\":\"split\",\"state\":\"frozen\",\"xSplit\":1,\"ySplit\":1,\"startRow\":1,\"startColumn\":1}",
                "{\"kind\":\"frozen\",\"state\":\"frozen\",\"xSplit\":1.5,\"ySplit\":0,\"startRow\":0,\"startColumn\":1}",
                "{\"kind\":\"frozen\",\"state\":\"frozen\",\"xSplit\":1,\"ySplit\":0,\"startRow\":0,\"startColumn\":16384}",
                "{\"kind\":\"split\",\"state\":\"split\",\"xSplit\":20,\"ySplit\":10,\"startRow\":0,\"startColumn\":0,\"activePane\":\"center\"}",
                "{\"kind\":\"frozen\",\"state\":\"frozen\",\"xSplit\":1,\"ySplit\":0,\"startRow\":0,\"startColumn\":1,\"referenceHint\":\"A1\"}")) {
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
                {"sheets":[{"id":"sheet-1","name":"Data","rowCount":10,"columnCount":2,"cells":{},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"protectionRules":[
                  {"id":"lock-outside-grid","scope":"range","range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":500,"endColumn":500},"locked":true,"allow":{}}
                ]}]}
                """);
        var mutation = withSortContext(new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":9,"startColumn":0,"endColumn":0},"sourceRows":[5,6,7,8,9,0,1,2,3,4]}
                """)), range(0, 9, 0, 0), "worksheet", null, false, 500);

        ServiceException error = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR));
        assertEquals("FORBIDDEN", error.code());

        var owner = registry.prepare(snapshot, mutation, WorkbookAclRole.OWNER);
        assertEquals(0, owner.affectedRanges().getFirst().startColumn());
        assertEquals(500, owner.affectedRanges().getFirst().endColumn());
        var updated = owner.descriptor().apply(snapshot, mutation);
        assertEquals(5, updated.path("sheets").get(0).path("protectionRules").get(0).path("range").path("startRow").asInt());
    }

    @Test
    void rowPermutationRemapsCrossSheetPivotAndSparklineSourcesWithoutMovingOwnerAnchors() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = crossSheetRowPermutationSnapshot(0, 0);
        OperationMutation mutation = withSortContext(new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":2,"startColumn":0,"endColumn":0},"sourceRows":[2,0,1]}
                """)), range(0, 2, 0, 0), "worksheet", null, false, 1);

        JsonNode current = registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, mutation);
        JsonNode owner = current.path("sheets").get(1);
        assertEquals(1, owner.path("sparklines").get(0).path("sourceRange").path("startRow").asInt());
        assertEquals(0, owner.path("sparklines").get(0).path("anchor").path("row").asInt());
        JsonNode pivot = owner.path("pivots").get(0);
        assertEquals(1, pivot.path("source").path("ranges").get(0).path("range").path("startRow").asInt());
        assertEquals(0, pivot.path("source").path("ranges").get(1).path("range").path("startRow").asInt());
        assertEquals(0, pivot.path("target").path("anchor").path("row").asInt());
        assertEquals(2, owner.path("pivots").get(1).path("source").path("range").path("startRow").asInt());
        JsonNode untouched = current.path("sheets").get(2);
        assertTrue(untouched.path("pivots").isMissingNode());
        assertTrue(untouched.path("sparklines").isMissingNode());
    }

    @Test
    void rowPermutationRejectsSplitCrossSheetPivotAndSparklineSourcesBeforeChangingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = crossSheetRowPermutationSnapshot(0, 1);
        OperationMutation mutation = withSortContext(new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":3,"startColumn":0,"endColumn":0},"sourceRows":[2,0,3,1]}
                """)), range(0, 3, 0, 0), "worksheet", null, false, 1);
        ObjectNode beforeSparklineRejection = snapshot.deepCopy();

        ServiceException sparklineError = assertThrows(ServiceException.class,
                () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, mutation));
        assertEquals("VALIDATION_ERROR", sparklineError.code());
        assertEquals(beforeSparklineRejection, snapshot);

        ((ArrayNode) snapshot.path("sheets").get(1).path("sparklines")).remove(0);
        ObjectNode beforePivotRejection = snapshot.deepCopy();
        ServiceException pivotError = assertThrows(ServiceException.class,
                () -> registry.prepare(snapshot, mutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, mutation));
        assertEquals("VALIDATION_ERROR", pivotError.code());
        assertEquals(beforePivotRejection, snapshot);
    }

    @Test
    void rowPermutationKeepsRuleAnchorsReportBindingsBandedRangesAndDrawingSourcesCanonical() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"definedNameModels":[],"cellStyleTemplates":[],"sheets":[
                  {"id":"sheet-1","name":"Data","rowCount":4,"columnCount":1,"cells":{"0":{"0":{"value":null,"formula":"=A1","formulaMetadata":{"kind":"normal","sourceFormula":"=A1"}}}},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"merges":[],
                   "conditionalFormats":[
                     {"id":"cf-implicit","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":2,"endColumn":2}],"type":"highlight","operator":"formula","value1":"=A1>0"},
                     {"id":"cf-explicit","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":3,"endColumn":3}],"formulaAnchor":{"sheetId":"sheet-1","row":0,"column":3},"type":"highlight","operator":"formula","value1":"=A1>0"},
                     {"id":"cf-literal","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":4,"endColumn":4}],"type":"highlight","operator":"greaterThan","value1":"0"}],
                   "dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],
                   "spillRanges":[],"protectionRules":[],"bandedRule":{"range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"firstColor":"#ffffff","secondColor":"#eeeeee"},
                   "reportSheet":{"templateSheetId":"sheet-1","bindings":[{"cell":{"row":0,"column":5},"expression":"field-id","kind":"field"}],"pagination":{"enabled":true,"repeatHeaderRows":[0]},"renderMode":"preview","layout":{"orientation":"portrait","marginTopPx":24,"marginRightPx":24,"marginBottomPx":24,"marginLeftPx":24},"dataEntry":[]}},
                  {"id":"sheet-2","name":"Drawing owner","rowCount":4,"columnCount":1,"cells":{},"drawingPayloads":{"camera-1":{"kind":"camera","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"refreshPolicy":"live"}}}
                ]}
                """);
        ObjectNode selected = range(0, 2, 0, 0);
        OperationMutation raw = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":2,"startColumn":0,"endColumn":0},"sourceRows":[2,0,1]}
                """));

        OperationMutation overstated = withSortContext(raw, selected, "worksheet", null, false, 6);
        ServiceException extentError = assertThrows(ServiceException.class,
                () -> registry.prepare(snapshot, overstated, WorkbookAclRole.OWNER));
        assertEquals("VALIDATION_ERROR", extentError.code());

        OperationMutation mutation = withSortContext(raw, selected, "worksheet", null, false, 5);
        var prepared = registry.prepare(snapshot, mutation, WorkbookAclRole.OWNER);
        var application = prepared.descriptor().applyWithPatch(snapshot, mutation);
        JsonNode updated = application.snapshot();
        assertTrue(application.structuralPatch() != null);
        assertEquals("rows.permuted", application.structuralPatch().mutationId());
        assertTrue(application.structuralPatch().formulaOwnerDeltas().stream().anyMatch(delta ->
                "formula-cell".equals(delta.kind())
                        && delta.beforeAddress().row() == 0
                        && delta.afterAddress().row() == 1
                        && "=A1".equals(delta.before().formula())
                        && "=A2".equals(delta.after().formula())));
        assertTrue(application.structuralPatch().formulaOwnerDeltas().stream().anyMatch(delta ->
                "formula-rule".equals(delta.kind()) && "cf-implicit".equals(delta.ruleId())
                        && "=A1>0".equals(delta.beforeFormula()) && "=A2>0".equals(delta.afterFormula())));
        JsonNode dataSheet = updated.path("sheets").get(0);
        assertEquals("=A2", dataSheet.path("cells").path("1").path("0").path("formula").asText());
        assertEquals(1, dataSheet.path("conditionalFormats").get(0).path("formulaAnchor").path("row").asInt());
        assertEquals("=A2>0", dataSheet.path("conditionalFormats").get(0).path("value1").asText());
        assertEquals(1, dataSheet.path("conditionalFormats").get(1).path("formulaAnchor").path("row").asInt());
        assertEquals("=A2>0", dataSheet.path("conditionalFormats").get(1).path("value1").asText());
        assertTrue(dataSheet.path("conditionalFormats").get(2).path("formulaAnchor").isMissingNode());
        assertEquals(1, dataSheet.path("bandedRule").path("range").path("startRow").asInt());
        assertEquals(1, dataSheet.path("reportSheet").path("bindings").get(0).path("cell").path("row").asInt());
        assertEquals(5, dataSheet.path("reportSheet").path("bindings").get(0).path("cell").path("column").asInt());
        assertEquals(1, updated.path("sheets").get(1).path("drawingPayloads").path("camera-1").path("sourceRange").path("startRow").asInt());
        assertEquals(2, updated.path("sheets").get(1).path("drawingPayloads").path("camera-1").path("sourceRange").path("endRow").asInt());
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
        filter = withDataRegionContext(filter, range(0, 4, 0, 2), "worksheet", null, false);
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
                {"schema":"WorkbookSnapshot","unitId":"book-1","name":"Book","dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":10,"cells":{}}],"printDocuments":[],"queryDefinitions":[]}
                """);
        OperationMutation print = new OperationMutation("pageLayout.paperSize.set", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","paperSize":"a4"}
                """));
        JsonNode current = registry.prepare(snapshot, print, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, print);
        assertEquals("a4", current.path("printDocuments").get(0).path("pageSetup").path("paperSize").asText());

        OperationMutation load = new OperationMutation("query.load.range", "sheet-1", mapper.readTree("""
                {"kind":"data-source-load","queryId":"query-1","queryDefinition":{"schema":"QueryDefinition","id":"query-1","name":"Block backed","connectorId":"json","connectorConfig":{"data":[]},"steps":[{"id":"trim-1","kind":"trim-text","name":"Trim","config":{"columns":["Name"]},"enabled":true},{"id":"split-1","kind":"split-column","name":"Split","config":{"column":"Name","delimiter":",","outputColumns":["First","Last"]},"enabled":true},{"id":"dedupe-1","kind":"remove-duplicates","name":"Dedupe","config":{"columns":["First"]},"enabled":true}],"sourceRevision":0},"target":{"kind":"range","sheetId":"sheet-1"},"sourceId":"query:query-1","source":null,"binding":null}
                """));
        current = registry.prepare(current, load, WorkbookAclRole.EDITOR).descriptor().apply(current, load);
        assertEquals("query-1", current.path("queryDefinitions").get(0).path("id").asText());
        assertEquals(0, current.path("sheets").get(0).path("cells").size());
    }

    @Test
    void pivotSparklineAndDrillDownReducersPreserveDefinitionsAndBlockBackedDetailMetadata() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},"sheets":[{"id":"sheet-1","name":"Sales","rowCount":20,"columnCount":10,"cells":{"0":{"0":{"value":"Region"},"1":{"value":"Amount"}},"1":{"0":{"value":"East"},"1":{"value":42}}},"pivots":[],"sparklines":[],"sparklineGroups":[]}]}
                """);
        OperationMutation pivot = new OperationMutation("pivot.add", "sheet-1", mapper.readTree("""
                {"schema":"PivotDefinition","id":"pivot-1","source":{"kind":"worksheet-range","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"target":{"sheetId":"sheet-1","anchor":{"row":4,"column":3}},"fieldCatalog":{"schema":"PivotFieldCatalog","fields":[{"fieldId":"sheet:sheet-1:column:0:range:0","name":"Region","dataType":"text","ordinal":0},{"fieldId":"sheet:sheet-1:column:1:range:0","name":"Amount","dataType":"number","ordinal":1}]},"layout":{"rows":[{"fieldId":"sheet:sheet-1:column:0:range:0","subtotal":{"mode":"automatic"}}],"columns":[],"filters":[{"kind":"manual","family":"manual","fieldId":"sheet:sheet-1:column:0:range:0","scope":"report","mode":"all","memberKeys":[]}],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[{"valueId":"value:amount","fieldId":"sheet:sheet-1:column:1:range:0","summarizeBy":"sum"}],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"},"refreshPolicy":{"mode":"on-change","preserveFormatting":true,"refreshOnLoad":true}}
                """));
        JsonNode current = registry.prepare(snapshot, pivot, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, pivot);
        assertEquals("pivot-1", current.path("sheets").get(0).path("pivots").get(0).path("id").asText());

        ObjectNode pivotWithDisplayOptions = (ObjectNode) pivot.params().deepCopy();
        ((ObjectNode) pivotWithDisplayOptions).set("presentation", mapper.readTree("""
                {"styleOptions":{"showRowHeaders":true,"showColumnHeaders":true,"showRowStripes":false,"showColumnStripes":false,"showLastColumn":false},"displayOptions":{"fillEmptyCells":false,"emptyCellText":"","showErrorValues":true,"errorCellText":"","showFieldHeaders":true,"autoFitColumnsOnUpdate":true}}
                """));
        JsonNode persistedPresentation = registry.prepare(snapshot, new OperationMutation("pivot.add", "sheet-1", pivotWithDisplayOptions), WorkbookAclRole.EDITOR)
                .descriptor().apply(snapshot, new OperationMutation("pivot.add", "sheet-1", pivotWithDisplayOptions));
        assertTrue(persistedPresentation.path("sheets").get(0).path("pivots").get(0).path("presentation").path("displayOptions").path("autoFitColumnsOnUpdate").asBoolean());

        ObjectNode invalidDisplayOptions = pivotWithDisplayOptions.deepCopy();
        ((ObjectNode) invalidDisplayOptions.path("presentation").path("displayOptions")).put("autoFitColumnsOnUpdate", "yes");
        ServiceException invalidDisplayOption = assertThrows(ServiceException.class, () -> registry.prepare(snapshot,
                new OperationMutation("pivot.add", "sheet-1", invalidDisplayOptions), WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", invalidDisplayOption.code());

        ObjectNode difference = (ObjectNode) pivot.params().deepCopy();
        ObjectNode differenceValue = (ObjectNode) difference.path("layout").path("values").get(0);
        differenceValue.set("showAs", mapper.createObjectNode()
                .put("kind", "difference")
                .put("baseFieldId", "sheet:sheet-1:column:0:range:0")
                .set("baseItem", mapper.createObjectNode().put("type", "text").put("value", "East")));
        registry.prepare(snapshot, new OperationMutation("pivot.add", "sheet-1", difference), WorkbookAclRole.EDITOR);

        ObjectNode missingOperand = (ObjectNode) difference.deepCopy();
        ((ObjectNode) missingOperand.path("layout").path("values").get(0).path("showAs")).remove("baseItem");
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, new OperationMutation("pivot.add", "sheet-1", missingOperand), WorkbookAclRole.EDITOR));

        ObjectNode highCardinality = (ObjectNode) pivot.params().deepCopy();
        ArrayNode members = mapper.createArrayNode();
        for (int index = 0; index < 10_001; index++) members.add("Member " + index);
        ((ObjectNode) highCardinality.path("fieldCatalog").path("fields").get(0)).set("values", members);
        OperationMutation highCardinalityMutation = new OperationMutation("pivot.add", "sheet-1", highCardinality);
        JsonNode highCardinalitySnapshot = registry.prepare(snapshot, highCardinalityMutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, highCardinalityMutation);
        assertEquals(10_001, highCardinalitySnapshot.path("sheets").get(0).path("pivots").get(0).path("fieldCatalog").path("fields").get(0).path("values").size());

        ObjectNode boundedManualFilter = (ObjectNode) pivot.params().deepCopy();
        ObjectNode manualFilter = (ObjectNode) boundedManualFilter.path("layout").path("filters").get(0);
        manualFilter.put("mode", "exclude");
        ArrayNode manualMembers = mapper.createArrayNode();
        for (int index = 0; index < 10_000; index++) {
            manualMembers.addObject().put("type", "text").put("value", "Member " + index);
        }
        manualFilter.set("memberKeys", manualMembers);
        registry.prepare(snapshot, new OperationMutation("pivot.add", "sheet-1", boundedManualFilter), WorkbookAclRole.EDITOR);

        manualMembers.addObject().put("type", "text").put("value", "Member 10000");
        ServiceException oversizedFilter = assertThrows(ServiceException.class, () -> registry.prepare(snapshot,
                new OperationMutation("pivot.add", "sheet-1", boundedManualFilter), WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", oversizedFilter.code());

        ObjectNode oversizedManualGroup = (ObjectNode) pivot.params().deepCopy();
        ObjectNode rowPlacement = (ObjectNode) oversizedManualGroup.path("layout").path("rows").get(0);
        ObjectNode group = rowPlacement.putObject("group");
        group.put("kind", "manual").putArray("groups").addObject()
                .put("groupId", "group-1").put("name", "Group 1").set("items", manualMembers);
        ServiceException oversizedGroup = assertThrows(ServiceException.class, () -> registry.prepare(snapshot,
                new OperationMutation("pivot.add", "sheet-1", oversizedManualGroup), WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", oversizedGroup.code());

        ObjectNode malformedValues = (ObjectNode) pivot.params().deepCopy();
        ((ObjectNode) malformedValues.path("fieldCatalog").path("fields").get(0)).set("values", mapper.createObjectNode());
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, new OperationMutation("pivot.add", "sheet-1", malformedValues), WorkbookAclRole.EDITOR));

        OperationMutation sparkline = new OperationMutation("sparkline.add", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sparkline":{"id":"spark-1","sheetId":"sheet-1","anchor":{"row":3,"column":0},"sourceRange":{"sheetId":"sheet-1","startRow":1,"endRow":1,"startColumn":1,"endColumn":1},"type":"line","color":"#2563eb"}}
                """));
        current = registry.prepare(current, sparkline, WorkbookAclRole.EDITOR).descriptor().apply(current, sparkline);
        assertEquals("spark-1", current.path("sheets").get(0).path("sparklines").get(0).path("id").asText());

        ObjectNode drillDownParams = (ObjectNode) mapper.readTree("""
                {"sheetId":"sheet-1","pivotId":"pivot-1","label":"East","sourceRowPaths":[{"sheetId":"sheet-1","row":1}],"targetSheetId":"detail-1","target":{"row":0,"column":0}}
                """);
        drillDownParams.set("detail", drillDownDetail("detail-1", "detail-source-1", 1, List.of("Region", "Amount")));
        OperationMutation drillDown = new OperationMutation("pivot.drilldown.add", "sheet-1", drillDownParams);
        current = registry.prepare(current, drillDown, WorkbookAclRole.EDITOR).descriptor().apply(current, drillDown);
        assertEquals("detail-1", current.path("sheets").get(1).path("id").asText());
        assertEquals("Region", current.path("sheets").get(1).path("cells").path("0").path("0").path("value").asText());
        assertEquals("detail-source-1", current.path("sheets").get(1).path("dataRegions").get(0).path("sourceId").asText());

        ObjectNode largeDrillDownParams = mapper.createObjectNode();
        largeDrillDownParams.put("sheetId", "sheet-1");
        largeDrillDownParams.put("pivotId", "pivot-1");
        largeDrillDownParams.put("label", "Large");
        ArrayNode largePaths = largeDrillDownParams.putArray("sourceRowPaths");
        for (int index = 0; index < 1_001; index++) {
            largePaths.addObject().put("sheetId", "sheet-1").put("row", 1).put("recordId", "record-" + index);
        }
        largeDrillDownParams.put("targetSheetId", "detail-large");
        largeDrillDownParams.putObject("target").put("row", 0).put("column", 0);
        largeDrillDownParams.set("detail", drillDownDetail("detail-large", "detail-source-large", 1_001, List.of("Region", "Amount")));
        OperationMutation largeDrillDown = new OperationMutation("pivot.drilldown.add", "sheet-1", largeDrillDownParams);
        current = registry.prepare(current, largeDrillDown, WorkbookAclRole.EDITOR).descriptor().apply(current, largeDrillDown);
        JsonNode largeDetail = current.path("sheets").get(2);
        assertEquals("worksheet", largeDetail.path("kind").asText());
        assertEquals("detail-large", largeDetail.path("id").asText());
        assertEquals(1_002, largeDetail.path("rowCount").asInt());
        assertEquals(26, largeDetail.path("columnCount").asInt());
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
                {"sheetId":"sheet-1","pivotId":"pivot-1","calculationProof":{"schema":"PivotCalculationProof","pivotId":"pivot-1","sourceRevision":"source-next","layoutRevision":"layout-next","filterRevision":"filter-next","occupiedRange":{"sheetId":"sheet-1","startRow":4,"endRow":8,"startColumn":3,"endColumn":5}},"previousCalculationProof":{"schema":"PivotCalculationProof","pivotId":"pivot-1","sourceRevision":"source-current","layoutRevision":"layout-current","filterRevision":"filter-current","occupiedRange":{"sheetId":"sheet-1","startRow":4,"endRow":4,"startColumn":3,"endColumn":3}},"layout":{"rows":[{"fieldId":"calculated:margin"}],"columns":[],"filters":[],"allowMultipleFiltersPerField":true,"collation":{"locale":"en-US","sensitivity":"variant","numeric":false,"caseFirst":"false"},"values":[{"valueId":"value:calculated:margin","fieldId":"calculated:margin","summarizeBy":"sum"}],"calculatedFields":[{"fieldId":"calculated:margin","name":"Margin","formula":"=amount*1.15"}],"calculatedItems":[{"fieldId":"calculated-item:amount:premium","targetFieldId":"sheet:sheet-1:column:1:range:0","name":"Premium","formula":"=amount*3"}],"subtotalLocation":"bottom","showRowGrandTotals":true,"showColumnGrandTotals":true,"reportLayout":"compact"}}
                """));

        JsonNode next = registry.applyPublicMutations(snapshot, List.of(update));
        JsonNode pivot = next.path("sheets").get(0).path("pivots").get(0);
        assertEquals("calculated:margin", pivot.path("layout").path("calculatedFields").get(0).path("fieldId").asText());
        assertEquals("calculated-item:amount:premium", pivot.path("layout").path("calculatedItems").get(0).path("fieldId").asText());
        assertEquals(2, pivot.path("fieldCatalog").path("fields").size());
        assertEquals(false, pivot.path("fieldCatalog").path("fields").toString().contains("calculated:margin"));

        ObjectNode invalidParams = (ObjectNode) update.params().deepCopy();
        ((ObjectNode) invalidParams.path("calculationProof").path("occupiedRange")).put("startRow", 5);
        JsonNode beforeInvalidProof = snapshot.deepCopy();
        ServiceException invalidProof = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(snapshot, List.of(new OperationMutation("pivot.update", "sheet-1", invalidParams))));
        assertEquals("VALIDATION_ERROR", invalidProof.code());
        assertEquals(beforeInvalidProof, snapshot);
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
                   {"kind":"top-items","family":"top-items","fieldId":"%s","scope":"field","valueId":"%s","direction":"top","mode":"items","threshold":1},
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

        ObjectNode validValueSortPivot = (ObjectNode) pivot.deepCopy();
        ((ObjectNode) validValueSortPivot.path("layout").path("rows").get(0)).set("sort", mapper.readTree("""
                {"direction":"descending","by":"value","valueId":"value:amount:count"}
                """));
        registry.prepare(snapshot, new OperationMutation("pivot.add", "sheet-1", validValueSortPivot), WorkbookAclRole.EDITOR);

        ObjectNode missingValueSortIdentity = (ObjectNode) validValueSortPivot.deepCopy();
        ((ObjectNode) missingValueSortIdentity.path("layout").path("rows").get(0).path("sort")).remove("valueId");
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot,
                new OperationMutation("pivot.add", "sheet-1", missingValueSortIdentity), WorkbookAclRole.EDITOR));

        ObjectNode labelSortWithValueIdentity = (ObjectNode) validValueSortPivot.deepCopy();
        ObjectNode labelSort = (ObjectNode) labelSortWithValueIdentity.path("layout").path("rows").get(0).path("sort");
        labelSort.put("by", "label");
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot,
                new OperationMutation("pivot.add", "sheet-1", labelSortWithValueIdentity), WorkbookAclRole.EDITOR));

        ObjectNode duplicateValuesPivot = (ObjectNode) pivot.deepCopy();
        ArrayNode duplicateValues = (ArrayNode) duplicateValuesPivot.path("layout").path("values");
        duplicateValues.add(duplicateValues.get(0).deepCopy());
        OperationMutation duplicateValuesMutation = new OperationMutation("pivot.add", "sheet-1", duplicateValuesPivot);
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, duplicateValuesMutation, WorkbookAclRole.EDITOR));

        ObjectNode legacyTopItems = (ObjectNode) pivot.deepCopy();
        ObjectNode legacyFilter = (ObjectNode) legacyTopItems.path("layout").path("filters").get(2);
        legacyFilter.remove("mode");
        legacyFilter.remove("threshold");
        legacyFilter.put("count", 1);
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot,
                new OperationMutation("pivot.add", "sheet-1", legacyTopItems), WorkbookAclRole.EDITOR));

        ObjectNode invalidTopMode = (ObjectNode) pivot.deepCopy();
        ObjectNode invalidModeFilter = (ObjectNode) invalidTopMode.path("layout").path("filters").get(2);
        invalidModeFilter.put("mode", "average");
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot,
                new OperationMutation("pivot.add", "sheet-1", invalidTopMode), WorkbookAclRole.EDITOR));

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

        ObjectNode missingValuePlacementPivot = (ObjectNode) pivot.deepCopy();
        ((ObjectNode) missingValuePlacementPivot.path("layout").path("filters").get(3)).remove("valueId");
        OperationMutation missingValuePlacement = new OperationMutation("pivot.add", "sheet-1", missingValuePlacementPivot);
        ServiceException missingValueError = assertThrows(ServiceException.class, () -> registry.prepare(snapshot, missingValuePlacement, WorkbookAclRole.EDITOR));
        assertEquals("VALIDATION_ERROR", missingValueError.code());
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
    void pivotWorksheetRangeSourcesEnforceAggregateFieldLimitBeforeMaterialization() throws Exception {
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":2,"columnCount":10001,"cells":{}}]}
                """);
        ObjectNode accepted = (ObjectNode) mapper.readTree("""
                {"kind":"worksheet-ranges","ranges":[
                  {"sourceId":"left","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":4999}},
                  {"sourceId":"right","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":5000,"endColumn":9999}}
                ]}
                """);

        PivotSourceResolver.Resolution resolution = PivotSourceResolver.resolve(snapshot, accepted);

        assertEquals(10_000, resolution.fieldIds().size());
        assertEquals("source:right:column:4999", resolution.fieldIds().get(9_999));

        ObjectNode oversized = accepted.deepCopy();
        ((ObjectNode) oversized.path("ranges").get(1).path("range")).put("endColumn", 10_000);
        ServiceException error = assertThrows(ServiceException.class,
                () -> PivotSourceResolver.resolve(snapshot, oversized));
        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals("Pivot worksheet range sources exceed the field limit", error.getMessage());
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

    private ObjectNode workbookSourceRangePermutationSnapshot() throws Exception {
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},
                 "sheets":[{"id":"sheet-1","name":"Data","rowCount":4,"columnCount":1,
                   "cells":{"0":{"0":{"value":"first"}},"1":{"0":{"value":"second"}},"2":{"0":{"value":"third"}},"3":{"0":{"value":"fourth"}}},
                   "pane":{"kind":"none"},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                   "merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                """);
        ObjectNode dataModel = (ObjectNode) snapshot.path("dataModel");
        ObjectNode table = ((ArrayNode) dataModel.path("tables")).addObject();
        table.put("id", "workbook-table").put("name", "Source rows").put("sourceSheetId", "sheet-1");
        table.putObject("sourceRange").put("sheetId", "sheet-1").put("startRow", 0).put("endRow", 2).put("startColumn", 0).put("endColumn", 0);
        table.put("rowCount", 2).put("blockSize", 128).put("revision", 0);
        table.putArray("fields").addObject().put("id", "value").put("name", "Value").put("ordinal", 0).put("type", "text");
        table.putArray("blocks");

        ObjectNode source = ((ArrayNode) dataModel.path("sources")).addObject();
        source.put("schema", "DataSourceManifest").put("version", 1).put("id", "data-source").put("name", "Source rows");
        source.put("kind", "chunked-table").put("sourceSheetId", "sheet-1");
        source.putObject("sourceRange").put("sheetId", "sheet-1").put("startRow", 0).put("endRow", 2).put("startColumn", 0).put("endColumn", 0);
        source.put("rowCount", 2).put("blockRowCount", 65_536).put("revision", 0);
        source.putArray("fields").addObject().put("id", "value").put("name", "Value").put("ordinal", 0).put("type", "text");
        source.putArray("blocks").addObject().put("id", "block-1").put("dataSourceId", "data-source")
                .put("startRow", 0).put("rowCount", 2).put("storageKey", "block-1")
                .put("checksum", "a".repeat(64)).put("byteLength", 1).put("encoding", "columnar-v1").put("revision", 0);
        return snapshot;
    }

    private ObjectNode crossSheetRowPermutationSnapshot(int sourceStartRow, int sourceEndRow) throws Exception {
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"definedNames":{},"definedNameModels":[],"sheets":[
                  {"id":"sheet-1","name":"Data","rowCount":4,"columnCount":2,"cells":{},"pane":{"kind":"none"},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]},
                  {"id":"sheet-2","name":"View","rowCount":4,"columnCount":3,"cells":{},"pane":{"kind":"none"},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]},
                  {"id":"sheet-3","name":"Untouched","rowCount":4,"columnCount":2,"cells":{}}
                ]}
                """);
        String source = """
                {"kind":"worksheet-ranges","ranges":[
                  {"sourceId":"source","range":{"sheetId":"sheet-1","startRow":%d,"endRow":%d,"startColumn":0,"endColumn":0}},
                  {"sourceId":"owner","range":{"sheetId":"sheet-2","startRow":0,"endRow":0,"startColumn":0,"endColumn":0}}
                ],"relationships":[]}
                """.formatted(sourceStartRow, sourceEndRow);
        ObjectNode owner = (ObjectNode) snapshot.path("sheets").get(1);
        ObjectNode pivot = (ObjectNode) pivotWithSource(source, "cross-sheet-pivot");
        ObjectNode target = (ObjectNode) pivot.get("target");
        target.put("sheetId", "sheet-2");
        ((ObjectNode) target.get("anchor")).put("row", 0).put("column", 2);
        ((ArrayNode) owner.path("pivots")).add(pivot);
        ObjectNode singlePivot = (ObjectNode) pivotWithSource(
                "{\"kind\":\"worksheet-range\",\"range\":{\"sheetId\":\"sheet-1\",\"startRow\":1,\"endRow\":1,\"startColumn\":0,\"endColumn\":0}}",
                "cross-sheet-single-pivot");
        ObjectNode singleTarget = (ObjectNode) singlePivot.get("target");
        singleTarget.put("sheetId", "sheet-2");
        ((ObjectNode) singleTarget.get("anchor")).put("row", 0).put("column", 2);
        ((ArrayNode) owner.path("pivots")).add(singlePivot);

        ObjectNode sparkline = mapper.createObjectNode();
        sparkline.put("id", "cross-sheet-sparkline").put("sheetId", "sheet-2").put("type", "line").put("color", "#000000");
        sparkline.putObject("anchor").put("row", 0).put("column", 1);
        sparkline.putObject("sourceRange").put("sheetId", "sheet-1").put("startRow", sourceStartRow).put("endRow", sourceEndRow).put("startColumn", 0).put("endColumn", 0);
        ((ArrayNode) owner.path("sparklines")).add(sparkline);
        return snapshot;
    }

    @Test
    void pivotRemovalFailsClosedWhenAChartStillReferencesThePivot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":20,"columnCount":10,"cells":{},"drawings":[
                  {"id":"pivot-chart-1","sheetId":"sheet-1","kind":"chart","payloadId":"pivot-chart-payload-1","anchor":{"kind":"absolute"},"transform":{"x":0,"y":0,"width":120,"height":80,"rotation":0},"zIndex":1}
                ],"drawingPayloads":{"pivot-chart-payload-1":{"kind":"chart","chartId":"pivot-chart-1","source":{"kind":"pivot","pivotId":"pivot-1"},"chartType":"column","subtype":"clustered","elements":{"hiddenData":"show"}}},"pivots":[
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
                {"sheets":[{"id":"sheet-1","name":"Sheet 1","rowCount":10,"columnCount":10,"cells":{},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"merges":[],"hiddenRows":[],"hiddenColumns":[],"rowHeightsPx":{},"columnWidthsPx":{},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"drawings":[],"drawingPayloads":{},"spillRanges":[],"sheetTables":[],"conditionalFormats":[],"dataValidations":[],"protectionRules":[],"outline":{"groups":[]},"sparklines":[],"pivots":[
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
    void axisProtectionUsesTheRequestedDimensionBandLikeTheClient() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":100,"columnCount":26,"cells":{},"protectionRules":[
                  {"id":"protected-column","scope":"range","sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":99,"startColumn":3,"endColumn":3},"locked":true,"allow":{}}
                ]}]}
                """);

        OperationMutation deleteOtherColumns = new OperationMutation("columns.deleted", "sheet-1", mapper.readTree(
                "{\"sheetId\":\"sheet-1\",\"at\":1,\"count\":2}"));
        var prepared = registry.prepare(snapshot, deleteOtherColumns, WorkbookAclRole.EDITOR);
        assertEquals(new RangeRef("sheet-1", 0, 99, 1, 2), prepared.affectedRanges().getFirst());

        OperationMutation deleteProtectedColumn = new OperationMutation("columns.deleted", "sheet-1", mapper.readTree(
                "{\"sheetId\":\"sheet-1\",\"at\":3,\"count\":1}"));
        ServiceException blocked = assertThrows(ServiceException.class,
                () -> registry.prepare(snapshot, deleteProtectedColumn, WorkbookAclRole.EDITOR));
        assertEquals("FORBIDDEN", blocked.code());

        OperationMutation insertRows = new OperationMutation("rows.inserted", "sheet-1", mapper.readTree(
                "{\"sheetId\":\"sheet-1\",\"at\":10,\"count\":2}"));
        assertEquals(new RangeRef("sheet-1", 10, 11, 0, 25), registry.resolveRanges(snapshot, insertRows).getFirst());
    }

    @Test
    void axisShiftRewritesAuxiliaryFormulaOwnersAndRejectsFormulaGroups() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,"cells":{},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"hiddenRows":[],"hiddenColumns":[],"rowHeightsPx":{},"columnWidthsPx":{},"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"spillRanges":[],"protectionRules":[]},
                  {"id":"sheet-2","name":"Other","rowCount":5,"columnCount":3,"cells":{},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}
                ]}
                """);
        ObjectNode firstRow = ((ObjectNode) snapshot.path("sheets").get(0).path("cells")).putObject("0");
        ObjectNode movedCell = firstRow.putObject("0");
        movedCell.putNull("value").put("formula", "=A1");
        movedCell.set("formulaMetadata", mapper.readTree("""
                {"kind":"normal","sourceFormula":"=A1"}
                """));
        movedCell.set("presentation", mapper.readTree("""
                {"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=A1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}
                """));
        ObjectNode auxiliaryOnlyCell = firstRow.putObject("1");
        auxiliaryOnlyCell.putNull("value");
        auxiliaryOnlyCell.set("formulaMetadata", mapper.readTree("""
                {"kind":"normal","sourceFormula":"=Sheet1!A1"}
                """));
        auxiliaryOnlyCell.set("presentation", mapper.readTree("""
                {"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=Sheet1!A1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}
                """));
        ObjectNode externalFormula = ((ObjectNode) snapshot.path("sheets").get(1).path("cells")).putObject("0").putObject("0");
        externalFormula.putNull("value").put("formula", "=Sheet1!A1");

        OperationMutation insert = new OperationMutation("rows.inserted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","at":0,"count":1}
                """));
        JsonNode shifted = registry.applyPublicMutations(snapshot, List.of(insert));
        JsonNode moved = shifted.path("sheets").get(0).path("cells").path("1").path("0");
        assertEquals("=A2", moved.path("formula").asText());
        assertEquals("=A2", moved.path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=A2", moved.path("presentation").path("source").path("formula").asText());
        JsonNode auxiliaryOnly = shifted.path("sheets").get(0).path("cells").path("1").path("1");
        assertEquals("=Sheet1!A2", auxiliaryOnly.path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=Sheet1!A2", auxiliaryOnly.path("presentation").path("source").path("formula").asText());
        assertEquals("=Sheet1!A2", shifted.path("sheets").get(1).path("cells").path("0").path("0").path("formula").asText());

        ObjectNode groupedBandSnapshot = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) groupedBandSnapshot.path("sheets").get(0).path("cells").path("0").path("0"))
                .set("formulaMetadata", mapper.readTree("""
                        {"kind":"shared","range":"A1:A2","sourceFormula":"=A1"}
                        """));
        JsonNode groupedBandBefore = groupedBandSnapshot.deepCopy();
        ServiceException groupedBandRejection = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(groupedBandSnapshot, List.of(insert)));
        assertEquals("SERVICE_UNAVAILABLE", groupedBandRejection.code());
        assertEquals(groupedBandBefore, groupedBandSnapshot);

        ObjectNode groupedDependentSnapshot = (ObjectNode) snapshot.deepCopy();
        ObjectNode groupedDependent = (ObjectNode) groupedDependentSnapshot.path("sheets").get(1).path("cells").path("0").path("0");
        groupedDependent.set("formulaMetadata", mapper.readTree("""
                {"kind":"shared","range":"A1:A2","sourceFormula":"=Sheet1!A1"}
                """));
        JsonNode groupedDependentBefore = groupedDependentSnapshot.deepCopy();
        ServiceException groupedDependentRejection = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(groupedDependentSnapshot, List.of(insert)));
        assertEquals("SERVICE_UNAVAILABLE", groupedDependentRejection.code());
        assertEquals(groupedDependentBefore, groupedDependentSnapshot);
    }

    @Test
    void structuralAxisRewriteCoversPersistedFormulaOwnersAndRejectsRemovedTemplateAnchors() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[
                  {"id":"formula-view","name":"Formula View","tableId":"source-table","fields":[
                    {"fieldId":"view-calc","caption":"View Calc","formula":"=Sheet1!A1"}]}]},
                 "definedNames":{},"definedNameModels":[],"printDocuments":[],
                 "cellStyleTemplates":[{"id":"formula-template","name":"Formula Template","style":{},
                   "dataValidation":{"type":"custom","formulaAnchor":{"sheetId":"sheet-1","row":0,"column":0},
                     "formula1":"=A1","formula2":"=B1","listSource":{"kind":"formula","formula":"=C1"}}}],
                 "sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":5,"cells":{},
                   "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                   "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                   "merges":[],"hiddenRows":[],"hiddenColumns":[],"rowHeightsPx":{},"columnWidthsPx":{},
                   "drawings":[],"drawingPayloads":{"shape":{"kind":"shape","propertyFormula":"=A1"},
                    "chart-1":{"kind":"chart","chartId":"chart-1","chartType":"combo","subtype":"custom-combo",
                      "source":{"kind":"worksheet-ranges","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0}]},
                      "elements":{"hiddenData":"show","titleText":{"linkedFormula":"=A1"},
                        "legend":{"visible":true,"position":"bottom","text":{"linkedFormula":"=B1"}},
                        "categoryAxis":{"id":"category","position":"bottom","titleText":{"linkedFormula":"=C1"}},
                        "valueAxis":{"id":"value","position":"left","titleText":{"linkedFormula":"=D1"}},
                        "secondaryCategoryAxis":{"id":"secondary-category","position":"top","titleText":{"linkedFormula":"=F1"}},
                        "secondaryValueAxis":{"id":"secondary-value","position":"right","titleText":{"linkedFormula":"=G1"}},
                        "dataTable":{"visible":true,"font":{"linkedFormula":"=E1"}}}}},
                   "tableSheet":{"viewId":"formula-view","columns":[{"fieldId":"table-calc","caption":"Table Calc","formula":"=A1"}],"grouping":[]},
                   "spillRanges":[],"sheetTables":[],"conditionalFormats":[],"dataValidations":[],
                   "protectionRules":[],"outline":{"groups":[]},"sparklines":[],"pivots":[],"dataRegions":[],"hyperlinks":[]}]}
                """);
        OperationMutation insert = new OperationMutation("rows.inserted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","at":0,"count":1}
                """));

        var preparedInsert = registry.prepare(snapshot, insert, WorkbookAclRole.OWNER);
        var axisApplication = preparedInsert.descriptor().applyWithPatch(snapshot, insert);
        JsonNode shifted = axisApplication.snapshot();
        JsonNode sheet = shifted.path("sheets").get(0);
        JsonNode validation = shifted.path("cellStyleTemplates").get(0).path("dataValidation");
        assertEquals("=A2", sheet.path("tableSheet").path("columns").get(0).path("formula").asText());
        assertEquals("=A2", sheet.path("drawingPayloads").path("shape").path("propertyFormula").asText());
        assertEquals("=A2", sheet.path("drawingPayloads").path("chart-1").path("elements").path("titleText").path("linkedFormula").asText());
        assertEquals("=B2", sheet.path("drawingPayloads").path("chart-1").path("elements").path("legend").path("text").path("linkedFormula").asText());
        assertEquals("=C2", sheet.path("drawingPayloads").path("chart-1").path("elements").path("categoryAxis").path("titleText").path("linkedFormula").asText());
        assertEquals("=D2", sheet.path("drawingPayloads").path("chart-1").path("elements").path("valueAxis").path("titleText").path("linkedFormula").asText());
        assertEquals("=F2", sheet.path("drawingPayloads").path("chart-1").path("elements").path("secondaryCategoryAxis").path("titleText").path("linkedFormula").asText());
        assertEquals("=G2", sheet.path("drawingPayloads").path("chart-1").path("elements").path("secondaryValueAxis").path("titleText").path("linkedFormula").asText());
        assertEquals("=E2", sheet.path("drawingPayloads").path("chart-1").path("elements").path("dataTable").path("font").path("linkedFormula").asText());
        var formulaObjectDeltas = axisApplication.structuralPatch().formulaOwnerDeltas().stream()
                .filter(delta -> "formula-object".equals(delta.kind())).toList();
        assertEquals(13, formulaObjectDeltas.size());
        var formulaObjectPatch = new StructuralPatch(StructuralPatch.VERSION, "rows.inserted", formulaObjectDeltas);
        JsonNode restoredObjects = registry.applyStructuralPatch(shifted, formulaObjectPatch.inverse("rows.deleted"));
        assertEquals("=A1", restoredObjects.path("sheets").get(0).path("tableSheet").path("columns").get(0).path("formula").asText());
        assertEquals("=A1", restoredObjects.path("sheets").get(0).path("drawingPayloads").path("shape").path("propertyFormula").asText());
        assertEquals("=Sheet1!A1", restoredObjects.path("dataModel").path("views").get(0).path("fields").get(0).path("formula").asText());
        assertEquals("=A1", restoredObjects.path("cellStyleTemplates").get(0).path("dataValidation").path("formula1").asText());
        assertEquals("=B1", restoredObjects.path("cellStyleTemplates").get(0).path("dataValidation").path("formula2").asText());
        assertEquals("=C1", restoredObjects.path("cellStyleTemplates").get(0).path("dataValidation").path("listSource").path("formula").asText());
        assertEquals("=A1", restoredObjects.path("sheets").get(0).path("drawingPayloads").path("chart-1")
                .path("elements").path("titleText").path("linkedFormula").asText());
        assertEquals("=Sheet1!A2", shifted.path("dataModel").path("views").get(0).path("fields").get(0).path("formula").asText());
        assertEquals(1, validation.path("formulaAnchor").path("row").asInt());
        assertEquals("=A2", validation.path("formula1").asText());
        assertEquals("=B2", validation.path("formula2").asText());
        assertEquals("=C2", validation.path("listSource").path("formula").asText());

        OperationMutation cellShift = new OperationMutation("cells.inserted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":1,"endRow":1,"startColumn":0,"endColumn":0},
                 "affectedBand":{"sheetId":"sheet-1","startRow":1,"endRow":5,"startColumn":0,"endColumn":0},"operation":"insert","axis":"row"}
                """));
        var preparedCellShift = registry.prepare(shifted, cellShift, WorkbookAclRole.OWNER);
        var cellShiftApplication = preparedCellShift.descriptor().applyWithPatch(shifted, cellShift);
        JsonNode cellShifted = cellShiftApplication.snapshot();
        assertEquals(5, cellShiftApplication.structuralPatch().formulaOwnerDeltas().stream()
                .filter(delta -> "formula-object".equals(delta.kind())).count());
        JsonNode shiftedSheet = cellShifted.path("sheets").get(0);
        JsonNode shiftedValidation = cellShifted.path("cellStyleTemplates").get(0).path("dataValidation");
        assertEquals("=A3", shiftedSheet.path("tableSheet").path("columns").get(0).path("formula").asText());
        assertEquals("=A3", shiftedSheet.path("drawingPayloads").path("shape").path("propertyFormula").asText());
        assertEquals("=A3", shiftedSheet.path("drawingPayloads").path("chart-1").path("elements").path("titleText").path("linkedFormula").asText());
        assertEquals("=Sheet1!A3", cellShifted.path("dataModel").path("views").get(0).path("fields").get(0).path("formula").asText());
        assertEquals(2, shiftedValidation.path("formulaAnchor").path("row").asInt());
        assertEquals("=A3", shiftedValidation.path("formula1").asText());

        OperationMutation move = new OperationMutation("range.move", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":2,"endRow":2,"startColumn":0,"endColumn":0},"targetOrigin":{"row":3,"column":1}}
                """));
        var preparedMove = registry.prepare(cellShifted, move, WorkbookAclRole.OWNER);
        var moveApplication = preparedMove.descriptor().applyWithPatch(cellShifted, move);
        JsonNode moved = moveApplication.snapshot();
        assertEquals(5, moveApplication.structuralPatch().formulaOwnerDeltas().stream()
                .filter(delta -> "formula-object".equals(delta.kind())).count());
        JsonNode movedSheet = moved.path("sheets").get(0);
        JsonNode movedValidation = moved.path("cellStyleTemplates").get(0).path("dataValidation");
        assertEquals("=B4", movedSheet.path("tableSheet").path("columns").get(0).path("formula").asText());
        assertEquals("=B4", movedSheet.path("drawingPayloads").path("shape").path("propertyFormula").asText());
        assertEquals("=B4", movedSheet.path("drawingPayloads").path("chart-1").path("elements").path("titleText").path("linkedFormula").asText());
        assertEquals("=Sheet1!B4", moved.path("dataModel").path("views").get(0).path("fields").get(0).path("formula").asText());
        assertEquals(3, movedValidation.path("formulaAnchor").path("row").asInt());
        assertEquals(1, movedValidation.path("formulaAnchor").path("column").asInt());
        assertEquals("=B4", movedValidation.path("formula1").asText());
        assertEquals("=B2", movedValidation.path("formula2").asText());
        assertEquals("=C2", movedValidation.path("listSource").path("formula").asText());

        ObjectNode anchoredSnapshot = (ObjectNode) snapshot.deepCopy();
        OperationMutation deleteAnchor = new OperationMutation("rows.deleted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","at":0,"count":1}
                """));
        JsonNode before = anchoredSnapshot.deepCopy();
        ServiceException rejected = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(anchoredSnapshot, List.of(deleteAnchor)));
        assertEquals("VALIDATION_ERROR", rejected.code());
        assertEquals(before, anchoredSnapshot);

        ObjectNode outOfBoundsAnchor = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) outOfBoundsAnchor.path("cellStyleTemplates").get(0).path("dataValidation").path("formulaAnchor"))
                .put("row", 1_048_576);
        ServiceException invalidCoordinate = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(outOfBoundsAnchor, List.of(insert)));
        assertEquals("VALIDATION_ERROR", invalidCoordinate.code());

        ObjectNode nullAnchor = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) nullAnchor.path("cellStyleTemplates").get(0).path("dataValidation")).putNull("formulaAnchor");
        ServiceException invalidNull = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(nullAnchor, List.of(insert)));
        assertEquals("VALIDATION_ERROR", invalidNull.code());
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
                {"definedNames":{"Sales":"=Sheet1!A2","OtherRelative":"=A2","LocalAnchor":"=A2"},"definedNameModels":[
                  {"name":"Sales","formula":"=Sheet1!A2","scope":"workbook"},
                  {"name":"OtherRelative","formula":"=A2","scope":"workbook","anchor":{"sheetId":"sheet-2","row":1,"column":0}},
                  {"name":"LocalAnchor","formula":"=A2","scope":"workbook","anchor":{"sheetId":"sheet-1","row":3,"column":0}},
                  {"name":"SheetScopedReference","formula":"=Sheet1!A2","scope":"sheet","sheetId":"sheet-2","anchor":{"sheetId":"sheet-2","row":1,"column":0}}],"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=A2"}},"1":{"0":{"value":10}}},"pane":{"kind":"frozen","state":"frozen","xSplit":0,"ySplit":1,"startRow":1,"startColumn":0},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"hiddenRows":[1],"rowHeightsPx":{"1":33},"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"spillRanges":[],"protectionRules":[]},
                  {"id":"sheet-2","name":"Other","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=Sheet1!A2"}}},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}
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
        assertEquals("=A2", current.path("definedNames").path("OtherRelative").asText());
        assertEquals("=A2", current.path("definedNameModels").get(1).path("formula").asText());
        assertEquals("=A3", current.path("definedNames").path("LocalAnchor").asText());
        assertEquals("=A3", current.path("definedNameModels").get(2).path("formula").asText());
        assertEquals(4, current.path("definedNameModels").get(2).path("anchor").path("row").asInt());
        assertEquals("=Sheet1!A3", current.path("definedNameModels").get(3).path("formula").asText());
        assertEquals(1, current.path("definedNameModels").get(3).path("anchor").path("row").asInt());
        assertEquals(2, current.path("sheets").get(0).path("hiddenRows").get(0).asInt());

        ObjectNode deletedAnchorSnapshot = (ObjectNode) snapshot.deepCopy();
        OperationMutation deleteAnchorRow = new OperationMutation("rows.deleted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","at":3,"count":1}
                """));
        ServiceException deletedAnchor = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(deletedAnchorSnapshot, List.of(deleteAnchorRow)));
        assertEquals("VALIDATION_ERROR", deletedAnchor.code());
        assertEquals(3, deletedAnchorSnapshot.path("definedNameModels").get(2).path("anchor").path("row").asInt());
    }

    @Test
    void cellShiftRejectsDataOwnersBeyondExtentAndPreservesUnrelatedSources() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        for (String axis : List.of("row", "column")) for (String ownerKind : List.of("data-region", "workbook-table", "data-source")) {
            ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                    {"definedNames":{},"definedNameModels":[],"dataModel":{"tables":[],"sources":[],"views":[],"relationships":[]},"sheets":[
                      {"id":"sheet-1","name":"Sheet1","rowCount":8,"columnCount":8,"cells":{"2":{"2":{"value":42}}},
                       "pane":{"kind":"none"},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                       "dataRegions":[],"sheetTables":[],"pivots":[]}]}
                    """);
            ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
            ObjectNode range = mapper.createObjectNode().put("sheetId", "sheet-1")
                    .put("startRow", "row".equals(axis) ? 20 : 2).put("endRow", "row".equals(axis) ? 20 : 2)
                    .put("startColumn", "column".equals(axis) ? 20 : 2).put("endColumn", "column".equals(axis) ? 20 : 2);
            ObjectNode region = null;
            if ("workbook-table".equals(ownerKind)) {
                ObjectNode table = ((ArrayNode) snapshot.path("dataModel").path("tables")).addObject()
                        .put("id", "far-table").put("name", "Far table").put("sourceSheetId", "sheet-1")
                        .put("rowCount", 0).put("blockSize", 128).put("revision", 0);
                table.set("sourceRange", range);
                table.putArray("fields");
                table.putArray("blocks");
            } else {
                ObjectNode source = ((ArrayNode) snapshot.path("dataModel").path("sources")).addObject()
                        .put("schema", "DataSourceManifest").put("version", 1).put("id", "far-source").put("name", "Far source")
                        .put("kind", "data-region".equals(ownerKind) ? "chunked-table" : "worksheet-range")
                        .put("rowCount", 0).put("blockRowCount", 65_536).put("revision", 0);
                source.putArray("fields").addObject().put("id", "f0").put("name", "Code").put("ordinal", 0).put("type", "text");
                source.putArray("blocks");
                if ("data-region".equals(ownerKind)) {
                    region = ((ArrayNode) sheet.get("dataRegions")).addObject().put("id", "far-region")
                            .put("sourceId", "far-source").put("headerRow", range.path("startRow").asInt()).put("revision", 0);
                    region.set("range", range);
                } else {
                    source.put("sourceSheetId", "sheet-1");
                    source.set("sourceRange", range);
                }
            }
            ObjectNode params = mapper.createObjectNode().put("sheetId", "sheet-1").put("operation", "insert").put("axis", axis);
            params.putObject("range").put("sheetId", "sheet-1").put("startRow", 2).put("endRow", 2).put("startColumn", 2).put("endColumn", 2);
            params.putObject("affectedBand").put("sheetId", "sheet-1").put("startRow", 2).put("startColumn", 2)
                    .put("endRow", "row".equals(axis) ? 7 : 2).put("endColumn", "column".equals(axis) ? 7 : 2);
            OperationMutation operation = new OperationMutation("cells.inserted", "sheet-1", params);
            JsonNode before = snapshot.deepCopy();
            ServiceException failure = assertThrows(ServiceException.class,
                    () -> registry.require(operation.id(), false).apply(snapshot, operation));
            assertEquals("SERVICE_UNAVAILABLE", failure.code());
            assertTrue(failure.getMessage().contains("cell shift intersects " + ownerKind.replace('-', ' ')));
            assertEquals(before, snapshot);

            range.put("startRow", 0).put("endRow", 0).put("startColumn", 0).put("endColumn", 0);
            if (region != null) region.put("headerRow", 0);
            JsonNode acceptedBefore = snapshot.deepCopy();
            JsonNode updated = registry.require(operation.id(), false).apply(snapshot, operation);
            assertEquals(acceptedBefore, snapshot);
            assertEquals(snapshot.path("dataModel"), updated.path("dataModel"));
            assertEquals(sheet.path("dataRegions"), updated.path("sheets").get(0).path("dataRegions"));
            assertEquals(42, updated.path("sheets").get(0).path("cells")
                    .path("row".equals(axis) ? "3" : "2").path("column".equals(axis) ? "3" : "2").path("value").asInt());
        }
    }

    @Test
    void cellInsertAndRowPermutationHaveDeterministicInverseFriendlySnapshots() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"definedNames":{"OtherRelative":"=A1","LocalAnchor":"=A2"},"definedNameModels":[
                  {"name":"OtherRelative","formula":"=A1","scope":"workbook","anchor":{"sheetId":"sheet-2","row":0,"column":0}},
                  {"name":"SheetScopedReference","formula":"=Sheet1!A1","scope":"sheet","sheetId":"sheet-2","anchor":{"sheetId":"sheet-2","row":0,"column":0}},
                  {"name":"LocalAnchor","formula":"=A2","scope":"workbook","anchor":{"sheetId":"sheet-1","row":1,"column":0}}],"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=A1","formulaValue":5,"formulaMetadata":{"kind":"normal","sourceFormula":"=A1"},"presentation":{"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=A1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}},"1":{"value":null,"formulaMetadata":{"kind":"normal","sourceFormula":"=Sheet1!A1"},"presentation":{"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=Sheet1!A1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}}},"1":{"0":{"value":"drop"}}},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"review":{"notesByCell":{"0:0":"n1"},"notesById":{"n1":{"id":"n1"}},"threadIdsByCell":{},"threadsById":{}},"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]},
                  {"id":"sheet-2","name":"Other","rowCount":5,"columnCount":3,"cells":{"0":{"0":{"value":null,"formula":"=A1","formulaValue":6}}}}]}
                """);
        OperationMutation shift = new OperationMutation("cells.inserted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"operation":"insert","axis":"row","affectedBand":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":0}}
                """));
        JsonNode current = registry.prepare(snapshot, shift, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, shift);
        assertEquals("=A2", current.path("sheets").get(0).path("cells").path("1").path("0").path("formula").asText());
        assertEquals("=A2", current.path("sheets").get(0).path("cells").path("1").path("0").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=A2", current.path("sheets").get(0).path("cells").path("1").path("0").path("presentation").path("source").path("formula").asText());
        assertEquals("=Sheet1!A2", current.path("sheets").get(0).path("cells").path("0").path("1").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=Sheet1!A2", current.path("sheets").get(0).path("cells").path("0").path("1").path("presentation").path("source").path("formula").asText());
        assertEquals("=A1", current.path("definedNames").path("OtherRelative").asText());
        assertEquals("=A1", current.path("definedNameModels").get(0).path("formula").asText());
        assertEquals("=Sheet1!A2", current.path("definedNameModels").get(1).path("formula").asText());
        assertEquals("=A3", current.path("definedNames").path("LocalAnchor").asText());
        assertEquals("=A3", current.path("definedNameModels").get(2).path("formula").asText());
        assertEquals(2, current.path("definedNameModels").get(2).path("anchor").path("row").asInt());
        assertEquals("n1", current.path("sheets").get(0).path("review").path("notesByCell").path("1:0").asText());

        ObjectNode groupedBandSnapshot = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) groupedBandSnapshot.path("sheets").get(0).path("cells").path("0").path("0"))
                .set("formulaMetadata", mapper.readTree("""
                        {"kind":"shared","range":"A1:A2","sourceFormula":"=A1"}
                        """));
        JsonNode groupedBandBefore = groupedBandSnapshot.deepCopy();
        ServiceException groupedBandRejection = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(groupedBandSnapshot, List.of(shift)));
        assertEquals("SERVICE_UNAVAILABLE", groupedBandRejection.code());
        assertEquals(groupedBandBefore, groupedBandSnapshot);

        ObjectNode groupedDependentSnapshot = (ObjectNode) snapshot.deepCopy();
        ObjectNode groupedDependent = (ObjectNode) groupedDependentSnapshot.path("sheets").get(0).path("cells").path("0").path("1");
        groupedDependent.put("formula", "=Sheet1!A1");
        groupedDependent.set("formulaMetadata", mapper.readTree("""
                {"kind":"shared","range":"B1:B2","sourceFormula":"=Sheet1!A1"}
                """));
        JsonNode groupedDependentBefore = groupedDependentSnapshot.deepCopy();
        ServiceException groupedDependentRejection = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(groupedDependentSnapshot, List.of(shift)));
        assertEquals("SERVICE_UNAVAILABLE", groupedDependentRejection.code());
        assertEquals(groupedDependentBefore, groupedDependentSnapshot);

        OperationMutation deleteAnchoredCell = new OperationMutation("cells.deleted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":2,"endRow":2,"startColumn":0,"endColumn":0},"operation":"delete","axis":"row","affectedBand":{"sheetId":"sheet-1","startRow":2,"endRow":4,"startColumn":0,"endColumn":0}}
                """));
        JsonNode beforeDeleteAnchoredCell = current;
        ServiceException deletedCellAnchor = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(beforeDeleteAnchoredCell, List.of(deleteAnchoredCell)));
        assertEquals("VALIDATION_ERROR", deletedCellAnchor.code());
        assertEquals(2, current.path("definedNameModels").get(2).path("anchor").path("row").asInt());

        OperationMutation restore = new OperationMutation("cells.inserted.restore", "sheet-1", mapper.readTree("""
                {"spec":{"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"operation":"insert","axis":"row","affectedBand":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":0}},"cells":[{"row":0,"column":0,"cell":{"value":null,"formula":"=A1","formulaValue":7,"formulaMetadata":{"kind":"normal","sourceFormula":"=A1"},"presentation":{"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=B1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}}},{"row":1,"column":0,"cell":{"value":"drop"}}]}
                """));
        current = registry.prepare(current, restore, WorkbookAclRole.EDITOR).descriptor().apply(current, restore);
        assertEquals("=A1", current.path("sheets").get(0).path("cells").path("0").path("0").path("formula").asText());
        assertEquals("drop", current.path("sheets").get(0).path("cells").path("1").path("0").path("value").asText());
        ((ObjectNode) current.path("sheets").get(1).path("cells").path("0").path("0")).put("formulaValue", 6);

        OperationMutation permutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceRows":[1,0]}
                """));
        permutation = withSortContext(permutation, range(0, 1, 0, 0), "worksheet", null, false, 2);
        current = registry.prepare(current, permutation, WorkbookAclRole.EDITOR).descriptor().apply(current, permutation);
        assertEquals("drop", current.path("sheets").get(0).path("cells").path("0").path("0").path("value").asText());
        JsonNode movedFormula = current.path("sheets").get(0).path("cells").path("1").path("0");
        assertEquals("=A2", movedFormula.path("formula").asText());
        assertEquals("=A2", movedFormula.path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=B2", movedFormula.path("presentation").path("source").path("formula").asText());
        assertTrue(movedFormula.path("formulaValue").isMissingNode());
        assertTrue(current.path("sheets").get(1).path("cells").path("0").path("0").path("formulaValue").isMissingNode());

        OperationMutation rawInversePermutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceRows":[1,0]}
                """));
        OperationMutation inversePermutation = withSortContext(rawInversePermutation, range(0, 1, 0, 0), "worksheet", null, false, 2);
        current = registry.prepare(current, inversePermutation, WorkbookAclRole.EDITOR).descriptor().apply(current, inversePermutation);
        JsonNode restoredFormula = current.path("sheets").get(0).path("cells").path("0").path("0");
        assertEquals("=A1", restoredFormula.path("formula").asText());
        assertEquals("=A1", restoredFormula.path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=B1", restoredFormula.path("presentation").path("source").path("formula").asText());
    }

    @Test
    void structuralAxisAndPermutationPatchesPreserveWholeAxisWorksheetNames() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":4,"columnCount":2,
                   "cells":{"0":{"0":{"value":null,"formula":"=SUM('Budget A1'!B:B)+A1"}},"1":{"0":{"value":"second"}}},
                   "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                   "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                   "merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],
                   "drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]},
                  {"id":"budget-id","name":"Budget A1","rowCount":4,"columnCount":2,"cells":{},
                   "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                   "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}]}
                """);
        JsonNode original = snapshot.deepCopy();
        OperationMutation insertion = new OperationMutation("rows.inserted", "sheet-1",
                mapper.readTree("""
                        {"sheetId":"sheet-1","at":0,"count":1}
                        """));
        MutationApplication inserted = registry.require("rows.inserted", false).applyWithPatch(snapshot, insertion);
        assertEquals("=SUM('Budget A1'!B:B)+A2", inserted.snapshot().path("sheets").get(0)
                .path("cells").path("1").path("0").path("formula").asText());
        assertEquals(1, inserted.structuralPatch().formulaOwnerDeltas().size());
        assertEquals("=SUM('Budget A1'!B:B)+A2", inserted.structuralPatch().formulaOwnerDeltas().get(0).after().formula());
        assertEquals(original, snapshot);

        OperationMutation rawPermutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceRows":[1,0]}
                """));
        OperationMutation permutation = withSortContext(rawPermutation, range(0, 1, 0, 0), "worksheet", null, false, 1);
        MutationApplication permuted = registry.require("rows.permuted", false).applyWithPatch(snapshot, permutation);
        assertEquals("=SUM('Budget A1'!B:B)+A2", permuted.snapshot().path("sheets").get(0)
                .path("cells").path("1").path("0").path("formula").asText());
        assertEquals(1, permuted.structuralPatch().formulaOwnerDeltas().size());
        assertEquals("=SUM('Budget A1'!B:B)+A2", permuted.structuralPatch().formulaOwnerDeltas().get(0).after().formula());
        JsonNode restored = registry.require("rows.permuted", false).apply(permuted.snapshot(), permutation);
        assertEquals("=SUM('Budget A1'!B:B)+A1", restored.path("sheets").get(0)
                .path("cells").path("0").path("0").path("formula").asText());
        assertEquals(original, snapshot);
    }

    @Test
    void rowPermutationRejectsFormulaGroupsAndUnsupportedReferenceKindsAtomically() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        for (String formula : List.of("=A1", "=[Book]Sheet1!A1", "=SUM(Sheet1!1:3)",
                "=SUM('[Book.xlsx]Budget A1'!B:B)", "=SUM('Budget A1'!1:3)")) {
            ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                    {"sheets":[
                      {"id":"sheet-1","name":"Sheet1","rowCount":4,"columnCount":2,
                       "cells":{"0":{"0":{"value":null,"formula":"=A1","formulaValue":7}},"1":{"0":{"value":"second"}}},
                       "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                       "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                       "merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],
                       "drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                    """);
            ObjectNode formulaCell = (ObjectNode) snapshot.path("sheets").get(0).path("cells").path("0").path("0");
            formulaCell.put("formula", formula);
            if ("=A1".equals(formula)) formulaCell.set("formulaMetadata", mapper.readTree("""
                    {"kind":"shared","range":"A1:A2","sourceFormula":"=A1"}
                    """));
            JsonNode before = snapshot.deepCopy();
            OperationMutation rawPermutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                    {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceRows":[1,0]}
                    """));
            OperationMutation permutation = withSortContext(rawPermutation, range(0, 1, 0, 0), "worksheet", null, false, 1);

            ServiceException rejection = assertThrows(ServiceException.class,
                    () -> registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation));
            assertEquals("SERVICE_UNAVAILABLE", rejection.code());
            assertEquals(before, snapshot);
        }
    }

    @Test
    void structuralDeletionClampsFrozenPaneAndViewportInsideDeletedInterval() throws Exception {
        StructuralMutationDescriptor descriptor = new StructuralMutationDescriptor("rows.deleted");
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":20,"columnCount":4,
                  "cells":{},"pane":{"kind":"frozen","state":"frozen","xSplit":0,"ySplit":7,"startRow":6,"startColumn":0},
                  "defaultRowHeightPx":20,"defaultColumnWidthPx":64,"hiddenRows":[],"hiddenColumns":[],
                  "rowHeightsPx":{},"columnWidthsPx":{},"merges":[],"conditionalFormats":[],"dataValidations":[],
                  "pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],
                  "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                  "spillRanges":[],"protectionRules":[],"outline":{"groups":[]}}]}
                """);
        OperationMutation mutation = new OperationMutation("rows.deleted", "sheet-1",
                mapper.readTree("{\"sheetId\":\"sheet-1\",\"at\":5,\"count\":3}"));

        JsonNode result = descriptor.applyWithPatch(snapshot, mutation).snapshot();

        assertEquals(5, result.path("sheets").get(0).path("pane").path("ySplit").intValue());
        assertEquals(5, result.path("sheets").get(0).path("pane").path("startRow").intValue());

        ObjectNode overflow = snapshot.deepCopy();
        ((ObjectNode) overflow.path("sheets").get(0).path("pane")).put("startRow", 1_048_575);
        JsonNode overflowBefore = overflow.deepCopy();
        OperationMutation insertion = new OperationMutation("rows.inserted", "sheet-1",
                mapper.readTree("{\"sheetId\":\"sheet-1\",\"at\":0,\"count\":1}"));

        ServiceException error = assertThrows(ServiceException.class,
                () -> new StructuralMutationDescriptor("rows.inserted").applyWithPatch(overflow, insertion));

        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals(overflowBefore, overflow);
    }

    @Test
    void rowPermutationRejectsFormulaOffsetsThatWouldMakeTheInverseIrreversible() throws Exception {
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":4,"columnCount":2,
                   "cells":{"0":{"0":{"value":"first"}},"1":{"0":{"value":null,"formula":"=A1","formulaValue":7}}},
                   "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                   "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                   "merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],
                   "drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                """);
        JsonNode before = snapshot.deepCopy();
        OperationMutation rawPermutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceRows":[1,0]}
                """));
        OperationMutation permutation = withSortContext(rawPermutation, range(0, 1, 0, 0), "worksheet", null, false, 1);

        ServiceException rejection = assertThrows(ServiceException.class,
                () -> new MutationDescriptorRegistry().prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation));
        assertEquals("SERVICE_UNAVAILABLE", rejection.code());
        assertEquals(before, snapshot);
    }

    @Test
    void rowInsertMovesCameraAndScreenshotSourceRangesAcrossDrawingPayloads() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[
                  {"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,"cells":{},"pane":{"kind":"none"},
                   "defaultRowHeightPx":20,"defaultColumnWidthPx":64,"merges":[],"conditionalFormats":[],"dataValidations":[],
                   "pivots":[],"sparklines":[],"drawings":[],"sheetTables":[],"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                   "spillRanges":[],"protectionRules":[],"drawingPayloads":{
                     "camera-1":{"kind":"camera","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1},"refreshPolicy":"live"},
                     "screenshot-1":{"kind":"screenshot","sourceRange":{"sheetId":"sheet-1","startRow":2,"endRow":3,"startColumn":0,"endColumn":1},"includeGridlines":true,"capturedAt":"2026-01-01T00:00:00Z"}}},
                  {"id":"sheet-2","name":"Other","rowCount":5,"columnCount":3,"cells":{}}
                ]}
                """);
        OperationMutation insert = new OperationMutation("rows.inserted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","at":0,"count":1}
                """));

        JsonNode current = registry.prepare(snapshot, insert, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, insert);
        JsonNode payloads = current.path("sheets").get(0).path("drawingPayloads");
        assertEquals(1, payloads.path("camera-1").path("sourceRange").path("startRow").asInt());
        assertEquals(2, payloads.path("camera-1").path("sourceRange").path("endRow").asInt());
        assertEquals(3, payloads.path("screenshot-1").path("sourceRange").path("startRow").asInt());
        assertEquals(4, payloads.path("screenshot-1").path("sourceRange").path("endRow").asInt());

        OperationMutation removeCameraSource = new OperationMutation("rows.deleted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","at":0,"count":2}
                """));
        ServiceException rejected = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(snapshot, List.of(removeCameraSource)));
        assertEquals("VALIDATION_ERROR", rejected.code());
    }

    @Test
    void duplicateSheetRewritesEverySheetScopedFormulaOwner() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"definedNameModels":[{"name":"LocalName","scope":"sheet","sheetId":"sheet-1","formula":"=Sheet1!A1","anchor":{"sheetId":"sheet-1","row":0,"column":0}}],
                 "sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":5,"columnCount":3,
                  "cells":{"0":{"0":{"formula":"=Sheet1!A1","formulaMetadata":{"kind":"normal","sourceFormula":"=Sheet1!A1"},"presentation":{"kind":"barcode","source":{"kind":"formula","formula":"=Sheet1!A1"}}}}},
                  "tableSheet":{"columns":[{"fieldId":"calculated","formula":"=Sheet1!A1"}]},
                  "merges":[],"conditionalFormats":[{"id":"cf-1","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0}],"type":"highlight","operator":"formula","value1":"=Sheet1!A1"}],
                  "dataValidations":[{"id":"dv-1","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0}],"type":"custom","formula1":"=Sheet1!A1","listSource":{"kind":"formula","formula":"=Sheet1!A1"}}],
                  "pivots":[{"id":"pivot-1"}],"sparklines":[],"sparklineGroups":[],"drawings":[],"drawingGroups":[],"drawingPayloads":{"shape-1":{"kind":"shape","propertyFormula":"=Sheet1!A1"},"screenshot-1":{"kind":"screenshot","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}},"timeline-1":{"kind":"timeline","pivotId":"pivot-1","connections":[{"pivotId":"pivot-1"}]}},"sheetTables":[],
                  "spillRanges":[],"protectionRules":[],"dataRegions":[],"hyperlinks":[],
                  "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}]}
                """);
        OperationMutation duplicate = new OperationMutation("sheet.duplicated", "sheet-1", mapper.readTree("""
                {"sourceSheetId":"sheet-1","newId":"sheet-2","newName":"Sheet1 Copy"}
                """));

        JsonNode current = registry.prepare(snapshot, duplicate, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, duplicate);
        JsonNode copy = current.path("sheets").get(1);
        assertEquals("=Sheet1!A1", snapshot.path("sheets").get(0).path("conditionalFormats").get(0).path("value1").asText());
        assertEquals("='Sheet1 Copy'!A1", copy.path("conditionalFormats").get(0).path("value1").asText());
        assertEquals("='Sheet1 Copy'!A1", copy.path("dataValidations").get(0).path("formula1").asText());
        assertEquals("='Sheet1 Copy'!A1", copy.path("dataValidations").get(0).path("listSource").path("formula").asText());
        assertEquals("='Sheet1 Copy'!A1", copy.path("cells").path("0").path("0").path("formula").asText());
        assertEquals("='Sheet1 Copy'!A1", copy.path("cells").path("0").path("0").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("='Sheet1 Copy'!A1", copy.path("cells").path("0").path("0").path("presentation").path("source").path("formula").asText());
        assertEquals("='Sheet1 Copy'!A1", copy.path("tableSheet").path("columns").get(0).path("formula").asText());
        assertEquals("='Sheet1 Copy'!A1", copy.path("drawingPayloads").path("shape-1::sheet-2").path("propertyFormula").asText());
        assertEquals("sheet-2", copy.path("drawingPayloads").path("screenshot-1::sheet-2").path("sourceRange").path("sheetId").asText());
        assertEquals("pivot-1::sheet-2", copy.path("drawingPayloads").path("timeline-1::sheet-2").path("pivotId").asText());
        assertEquals("pivot-1::sheet-2", copy.path("drawingPayloads").path("timeline-1::sheet-2").path("connections").get(0).path("pivotId").asText());
        assertEquals("sheet-2", current.path("definedNameModels").get(1).path("anchor").path("sheetId").asText());
    }

    @Test
    void rangeMoveRewritesFormulaOwnersAndRejectsOverlappingDestinations() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},"definedNames":{},"definedNameModels":[],"printDocuments":[],
                 "sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":6,"columnCount":6,
                   "cells":{"0":{"0":{"value":7},"1":{"formula":"=A1","formulaMetadata":{"kind":"normal","sourceFormula":"=A1"},"presentation":{"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=A1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}},"4":{"formula":"=A1","formulaMetadata":{"kind":"normal","sourceFormula":"=A1"},"presentation":{"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=A1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}},"5":{"value":null,"formulaMetadata":{"kind":"normal","sourceFormula":"=A1"},"presentation":{"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=A1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}}},"2":{"3":{"value":"stale"}}},
                   "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                   "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                   "merges":[],"hiddenRows":[],"hiddenColumns":[],"rowHeightsPx":{},"columnWidthsPx":{},
                   "drawings":[],"drawingPayloads":{},"spillRanges":[],"sheetTables":[],"conditionalFormats":[],"dataValidations":[],
                   "protectionRules":[],"outline":{"groups":[]},"sparklines":[],"pivots":[],"dataRegions":[],"hyperlinks":[]}]}
                """);
        OperationMutation move = new OperationMutation("range.move", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":1},"targetOrigin":{"row":2,"column":2}}
                """));

        var prepared = registry.prepare(snapshot, move, WorkbookAclRole.EDITOR);
        assertEquals(2, prepared.affectedRanges().size());
        JsonNode moved = registry.applyPublicMutations(snapshot, List.of(move));
        JsonNode sheet = moved.path("sheets").get(0);
        assertEquals(7, sheet.path("cells").path("2").path("2").path("value").asInt());
        assertEquals("=C3", sheet.path("cells").path("2").path("3").path("formula").asText());
        assertEquals("=C3", sheet.path("cells").path("2").path("3").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=C3", sheet.path("cells").path("2").path("3").path("presentation").path("source").path("formula").asText());
        assertEquals("=C3", sheet.path("cells").path("0").path("4").path("formula").asText());
        assertEquals("=C3", sheet.path("cells").path("0").path("4").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=C3", sheet.path("cells").path("0").path("4").path("presentation").path("source").path("formula").asText());
        assertEquals("=C3", sheet.path("cells").path("0").path("5").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("=C3", sheet.path("cells").path("0").path("5").path("presentation").path("source").path("formula").asText());
        assertTrue(sheet.path("cells").path("2").path("3").path("value").isMissingNode());
        assertEquals(7, snapshot.path("sheets").get(0).path("cells").path("0").path("0").path("value").asInt());

        OperationMutation overlap = new OperationMutation("range.move", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"targetOrigin":{"row":0,"column":0}}
                """));
        ServiceException rejected = assertThrows(ServiceException.class, () -> registry.applyPublicMutations(snapshot, List.of(overlap)));
        assertEquals("VALIDATION_ERROR", rejected.code());

        ObjectNode groupedSourceSnapshot = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) groupedSourceSnapshot.path("sheets").get(0).path("cells").path("0").path("0"))
                .set("formulaMetadata", mapper.readTree("""
                        {"kind":"shared","range":"A1:A2","sourceFormula":"=1"}
                        """));
        JsonNode groupedSourceBefore = groupedSourceSnapshot.deepCopy();
        OperationMutation singleCellMove = new OperationMutation("range.move", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"targetOrigin":{"row":2,"column":2}}
                """));
        ServiceException groupedSourceRejection = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(groupedSourceSnapshot, List.of(singleCellMove)));
        assertEquals("SERVICE_UNAVAILABLE", groupedSourceRejection.code());
        assertEquals(groupedSourceBefore, groupedSourceSnapshot);

        ObjectNode groupedTargetSnapshot = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) groupedTargetSnapshot.path("sheets").get(0).path("cells").path("2"))
                .putObject("2").put("formula", "=1").set("formulaMetadata", mapper.readTree("""
                        {"kind":"shared","range":"C3:C4","sourceFormula":"=1"}
                        """));
        JsonNode groupedTargetBefore = groupedTargetSnapshot.deepCopy();
        ServiceException groupedTargetRejection = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(groupedTargetSnapshot, List.of(singleCellMove)));
        assertEquals("SERVICE_UNAVAILABLE", groupedTargetRejection.code());
        assertEquals(groupedTargetBefore, groupedTargetSnapshot);

        ObjectNode groupedDependentSnapshot = (ObjectNode) snapshot.deepCopy();
        ObjectNode groupedDependent = ((ObjectNode) groupedDependentSnapshot.path("sheets").get(0).path("cells").path("0"))
                .putObject("5");
        groupedDependent.putNull("value");
        groupedDependent.set("formulaMetadata", mapper.readTree("""
                {"kind":"shared","range":"F1:F2","sourceFormula":"=A1"}
                """));
        groupedDependent.set("presentation", mapper.readTree("""
                {"kind":"barcode","symbology":"qr","source":{"kind":"formula","formula":"=A1"},"parameters":{"symbology":"qr"},"options":{"foreground":"#000000","background":"#ffffff","showText":false,"labelPosition":"none","quietZone":0}}
                """));
        JsonNode groupedDependentBefore = groupedDependentSnapshot.deepCopy();
        ServiceException groupedDependentRejection = assertThrows(ServiceException.class,
                () -> registry.applyPublicMutations(groupedDependentSnapshot, List.of(singleCellMove)));
        assertEquals("SERVICE_UNAVAILABLE", groupedDependentRejection.code());
        assertEquals(groupedDependentBefore, groupedDependentSnapshot);

        ObjectNode partialRangeSnapshot = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) partialRangeSnapshot.path("sheets").get(0).path("cells").path("0"))
                .putObject("5").put("formula", "=SUM(A1:A4)");
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(partialRangeSnapshot, List.of(move)));

        ObjectNode wholeColumnSnapshot = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) wholeColumnSnapshot.path("sheets").get(0).path("cells").path("0"))
                .putObject("5").put("formula", "=SUM(A:A)");
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(wholeColumnSnapshot, List.of(move)));

        ObjectNode wholeRowSnapshot = (ObjectNode) snapshot.deepCopy();
        ((ObjectNode) wholeRowSnapshot.path("sheets").get(0).path("cells").path("0"))
                .putObject("5").put("formula", "=SUM(1:1)");
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(wholeRowSnapshot, List.of(move)));

        ObjectNode threeDimensionalSnapshot = (ObjectNode) snapshot.deepCopy();
        ObjectNode firstSheet = (ObjectNode) threeDimensionalSnapshot.path("sheets").get(0);
        ((ObjectNode) firstSheet.path("cells").path("0")).putObject("5").put("formula", "=SUM(Sheet1:Sheet2!A1)");
        ObjectNode secondSheet = firstSheet.deepCopy();
        secondSheet.put("id", "sheet-2");
        secondSheet.put("name", "Sheet2");
        ((ArrayNode) threeDimensionalSnapshot.path("sheets")).add(secondSheet);
        assertThrows(ServiceException.class, () -> registry.applyPublicMutations(threeDimensionalSnapshot, List.of(move)));
    }

    @Test
    void rangeMovePreservesGlobalNameReferencesRelativeToAnotherSheet() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},
                 "definedNames":{"Relative":"=A1"},
                 "definedNameModels":[{"name":"Relative","formula":"=A1","scope":"workbook","anchor":{"sheetId":"sheet-2","row":0,"column":0}}],
                 "sheets":[
                   {"id":"sheet-1","name":"Sheet1","rowCount":6,"columnCount":6,"cells":{"0":{"0":{"value":7}}},"pane":{"kind":"none"},"merges":[],"hiddenRows":[],"hiddenColumns":[],"rowHeightsPx":{},"columnWidthsPx":{},"drawings":[],"drawingPayloads":{},"spillRanges":[],"sheetTables":[],"conditionalFormats":[],"dataValidations":[],"protectionRules":[],"outline":{"groups":[]},"sparklines":[],"pivots":[],"dataRegions":[],"hyperlinks":[],"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}},
                   {"id":"sheet-2","name":"Sheet2","rowCount":6,"columnCount":6,"cells":{},"pane":{"kind":"none"},"merges":[],"hiddenRows":[],"hiddenColumns":[],"rowHeightsPx":{},"columnWidthsPx":{},"drawings":[],"drawingPayloads":{},"spillRanges":[],"sheetTables":[],"conditionalFormats":[],"dataValidations":[],"protectionRules":[],"outline":{"groups":[]},"sparklines":[],"pivots":[],"dataRegions":[],"hyperlinks":[],"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}}]}
                """
        );
        OperationMutation move = new OperationMutation("range.move", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},"targetOrigin":{"row":2,"column":2}}
                """));

        JsonNode moved = registry.applyPublicMutations(snapshot, List.of(move));

        assertEquals("=A1", moved.path("definedNames").path("Relative").asText());
        assertEquals("=A1", moved.path("definedNameModels").get(0).path("formula").asText());
    }

    @Test
    void fillSeriesUsesOneCanonicalTargetBandAndRejectsStaleBeforeImages() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","rowCount":6,"columnCount":3,
                  "cells":{"0":{"0":{"value":1}},"1":{"0":{"value":3}}},
                  "protectionRules":[]}]}
                """);
        OperationMutation fill = new OperationMutation("fill.applied", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},
                 "targetRange":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":0},
                 "direction":"down","mode":"series","writes":[
                   {"row":2,"column":0,"before":null,"after":{"value":5}},
                   {"row":3,"column":0,"before":null,"after":{"value":7}},
                   {"row":4,"column":0,"before":null,"after":{"value":9}}
                 ]}
                """));

        var prepared = registry.prepare(snapshot, fill, WorkbookAclRole.EDITOR);
        assertEquals(0, prepared.affectedRanges().get(0).startRow());
        assertEquals(4, prepared.affectedRanges().get(0).endRow());
        JsonNode next = prepared.descriptor().apply(snapshot, fill);
        assertEquals(5, next.path("sheets").get(0).path("cells").path("2").path("0").path("value").asInt());
        assertEquals(9, next.path("sheets").get(0).path("cells").path("4").path("0").path("value").asInt());
        assertEquals(2, snapshot.path("sheets").get(0).path("cells").size());

        OperationMutation stale = new OperationMutation("fill.applied", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","sourceRange":{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":0,"endColumn":0},
                 "targetRange":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},
                 "direction":"down","mode":"copy","writes":[
                   {"row":1,"column":0,"before":{"value":99},"after":{"value":1}}
                 ]}
                """));
        ServiceException conflict = assertThrows(ServiceException.class, () -> registry.prepare(next, stale, WorkbookAclRole.EDITOR).descriptor().apply(next, stale));
        assertEquals("CONFLICT", conflict.code());
    }

    @Test
    void rowPermutationMovesOnlyExactCellOwnersAndSplitsRangeMetadata() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":8,"columnCount":8,"cells":{"0":{"0":{"value":"a"}},"1":{"0":{"value":"b"}}},"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{"0:5":["outside-comment"]},"threadsById":{"outside-comment":{"id":"outside-comment","sheetId":"sheet-1","row":0,"column":5,"replies":[]}}},"hyperlinks":[{"row":0,"column":1,"hyperlink":{"id":"inside-link"}},{"row":0,"column":5,"hyperlink":{"id":"outside-link"}}],"merges":[],"conditionalFormats":[{"id":"cf-1","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":1}],"type":"highlight"}],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[{"id":"inside-drawing","sheetId":"sheet-1","kind":"shape","anchor":{"kind":"one-cell","row":0,"column":1}},{"id":"outside-drawing","sheetId":"sheet-1","kind":"shape","anchor":{"kind":"one-cell","row":0,"column":5}}],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                """);
        OperationMutation permutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":3,"startColumn":0,"endColumn":1},"sourceRows":[2,0,3,1]}
                """));
        permutation = withSortContext(permutation, range(0, 3, 0, 1), "worksheet", null, true, 7);

        JsonNode current = registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation);
        JsonNode sheet = current.path("sheets").get(0);
        assertEquals(true, sheet.path("cells").path("0").isMissingNode());
        assertEquals("b", sheet.path("cells").path("3").path("0").path("value").asText());
        assertEquals(1, sheet.path("hyperlinks").get(0).path("row").asInt());
        assertEquals(0, sheet.path("hyperlinks").get(1).path("row").asInt());
        assertEquals(0, sheet.path("review").path("threadsById").path("outside-comment").path("row").asInt());
        assertEquals(1, sheet.path("drawings").get(0).path("anchor").path("row").asInt());
        assertEquals(0, sheet.path("drawings").get(1).path("anchor").path("row").asInt());
        assertEquals(2, sheet.path("conditionalFormats").get(0).path("ranges").size());
        assertEquals(1, sheet.path("conditionalFormats").get(0).path("ranges").get(0).path("startRow").asInt());
        assertEquals(3, sheet.path("conditionalFormats").get(0).path("ranges").get(1).path("startRow").asInt());
    }

    @Test
    void rowPermutationRemapsWorkbookTableAndDataSourceRanges() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = workbookSourceRangePermutationSnapshot();
        OperationMutation raw = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":3,"startColumn":0,"endColumn":0},"sourceRows":[3,0,1,2]}
                """));
        OperationMutation permutation = withSortContext(raw, range(0, 3, 0, 0), "worksheet", null, false, 0);

        JsonNode current = registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation);

        assertEquals(1, current.path("dataModel").path("tables").get(0).path("sourceRange").path("startRow").asInt());
        assertEquals(3, current.path("dataModel").path("tables").get(0).path("sourceRange").path("endRow").asInt());
        assertEquals(1, current.path("dataModel").path("sources").get(0).path("sourceRange").path("startRow").asInt());
        assertEquals(3, current.path("dataModel").path("sources").get(0).path("sourceRange").path("endRow").asInt());
        assertEquals("first", current.path("sheets").get(0).path("cells").path("1").path("0").path("value").asText());
        assertEquals("third", current.path("sheets").get(0).path("cells").path("3").path("0").path("value").asText());
    }

    @Test
    void rowPermutationRejectsSplitWorkbookTableAndDataSourceRangesBeforeChangingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        OperationMutation raw = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":3,"startColumn":0,"endColumn":0},"sourceRows":[2,0,3,1]}
                """));
        OperationMutation permutation = withSortContext(raw, range(0, 3, 0, 0), "worksheet", null, false, 0);

        for (String ownerKind : List.of("workbook table", "data source")) {
            ObjectNode snapshot = workbookSourceRangePermutationSnapshot();
            ObjectNode dataModel = (ObjectNode) snapshot.path("dataModel");
            if (ownerKind.equals("workbook table")) ((ArrayNode) dataModel.path("sources")).removeAll();
            else ((ArrayNode) dataModel.path("tables")).removeAll();
            JsonNode before = snapshot.deepCopy();

            ServiceException error = assertThrows(ServiceException.class,
                    () -> registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation));

            assertEquals("VALIDATION_ERROR", error.code());
            assertEquals(before, snapshot);
        }
    }

    @Test
    void rowPermutationRewritesAnchoredRuleAndWorkbookFormulaOwners() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"definedNames":{"RelativeOwner":"=A1","OtherSheetOwner":"=A1"},
                 "definedNameModels":[{"name":"RelativeOwner","formula":"=A1","scope":"workbook","anchor":{"sheetId":"sheet-1","row":0,"column":6}},{"name":"OtherSheetOwner","formula":"=A1","scope":"workbook","anchor":{"sheetId":"sheet-2","row":0,"column":9}}],
                 "cellStyleTemplates":[{"id":"template-anchored","name":"Anchored validation","style":{},"dataValidation":{"type":"custom","formula1":"=A1>0","formula2":"=B1","listSource":{"kind":"formula","formula":"=C1:C2"},"formulaAnchor":{"sheetId":"sheet-1","row":0,"column":7}}},{"id":"template-other-sheet","name":"Other sheet validation","style":{},"dataValidation":{"type":"custom","formula1":"=A1>0","formulaAnchor":{"sheetId":"sheet-2","row":0,"column":10}}}],
                 "sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":4,"columnCount":2,
                   "cells":{"0":{"0":{"value":"first"}},"1":{"0":{"value":"second"}}},"pane":{"kind":"none"},
                   "defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                   "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                   "merges":[],"conditionalFormats":[{"id":"cf-anchored","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":4,"endColumn":4}],"formulaAnchor":{"sheetId":"sheet-1","row":0,"column":4},"type":"highlight","operator":"formula","value1":"=A1>0"},{"id":"cf-implicit-anchor","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":8,"endColumn":8}],"type":"highlight","operator":"formula","value1":"=A1>0"}],
                   "dataValidations":[{"id":"dv-anchored","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":0,"endRow":0,"startColumn":5,"endColumn":5}],"formulaAnchor":{"sheetId":"sheet-1","row":0,"column":5},"type":"custom","formula1":"=A1>0","formula2":"=B1","listSource":{"kind":"formula","formula":"=C1:C2"}}],
                   "pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]},
                  {"id":"sheet-2","name":"Other sheet","rowCount":4,"columnCount":2,"cells":{},"pane":{"kind":"none"},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]
                }
                """);
        OperationMutation rawPermutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceRows":[1,0]}
                """));
        OperationMutation permutation = withSortContext(rawPermutation, range(0, 1, 0, 0), "worksheet", null, false, 8);

        var preparation = registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR);
        assertEquals(8, preparation.affectedRanges().getFirst().endColumn());
        JsonNode current = preparation.descriptor().apply(snapshot, permutation);
        JsonNode sheet = current.path("sheets").get(0);
        assertEquals("=A2", current.path("definedNames").path("RelativeOwner").asText());
        assertEquals("=A2", current.path("definedNameModels").get(0).path("formula").asText());
        assertEquals(1, current.path("definedNameModels").get(0).path("anchor").path("row").asInt());
        assertEquals("=A1", current.path("definedNames").path("OtherSheetOwner").asText());
        assertEquals("=A1", current.path("definedNameModels").get(1).path("formula").asText());
        assertEquals(0, current.path("definedNameModels").get(1).path("anchor").path("row").asInt());
        assertEquals(1, sheet.path("conditionalFormats").get(0).path("formulaAnchor").path("row").asInt());
        assertEquals("=A2>0", sheet.path("conditionalFormats").get(0).path("value1").asText());
        assertEquals(1, sheet.path("conditionalFormats").get(1).path("formulaAnchor").path("row").asInt());
        assertEquals("=A2>0", sheet.path("conditionalFormats").get(1).path("value1").asText());
        assertEquals(1, sheet.path("dataValidations").get(0).path("formulaAnchor").path("row").asInt());
        assertEquals("=A2>0", sheet.path("dataValidations").get(0).path("formula1").asText());
        assertEquals("=B2", sheet.path("dataValidations").get(0).path("formula2").asText());
        assertEquals("=C2:C3", sheet.path("dataValidations").get(0).path("listSource").path("formula").asText());
        JsonNode templateValidation = current.path("cellStyleTemplates").get(0).path("dataValidation");
        assertEquals(1, templateValidation.path("formulaAnchor").path("row").asInt());
        assertEquals("=A2>0", templateValidation.path("formula1").asText());
        assertEquals("=B2", templateValidation.path("formula2").asText());
        assertEquals("=C2:C3", templateValidation.path("listSource").path("formula").asText());
        JsonNode otherSheetTemplateValidation = current.path("cellStyleTemplates").get(1).path("dataValidation");
        assertEquals(0, otherSheetTemplateValidation.path("formulaAnchor").path("row").asInt());
        assertEquals("=A1>0", otherSheetTemplateValidation.path("formula1").asText());
    }

    @Test
    void rowPermutationPreservesImplicitRuleAnchorWhenExactRangeFragmentsReorder() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":8,"columnCount":4,"cells":{},"pane":{"kind":"none"},
                  "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"merges":[],
                  "conditionalFormats":[{"id":"cf-implicit-fixed-anchor","sheetId":"sheet-1","ranges":[{"sheetId":"sheet-1","startRow":2,"endRow":6,"startColumn":3,"endColumn":3}],"type":"highlight","operator":"formula","value1":"=A1>0"}],
                  "dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                """);
        OperationMutation rawPermutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":4,"startColumn":0,"endColumn":1},"sourceRows":[3,0,2,1,4]}
                """));
        OperationMutation permutation = withSortContext(rawPermutation, range(0, 4, 0, 1), "worksheet", null, false, 3);

        JsonNode current = registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation);
        JsonNode rule = current.path("sheets").get(0).path("conditionalFormats").get(0);

        assertEquals(2, rule.path("formulaAnchor").path("row").asInt());
        assertEquals("=A1>0", rule.path("value1").asText());
    }

    @Test
    void rowPermutationRejectsInvalidAnchoredNameFormulaWithoutChangingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"definedNames":{"OutOfBoundsOwner":"=A1"},
                 "definedNameModels":[{"name":"OutOfBoundsOwner","formula":"=A1","scope":"workbook","anchor":{"sheetId":"sheet-1","row":1,"column":8}}],
                 "sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":4,"columnCount":2,
                   "cells":{"0":{"0":{"value":"first"}},"1":{"0":{"value":"second"}}},"pane":{"kind":"none"},
                   "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                   "merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                """);
        JsonNode before = snapshot.deepCopy();
        OperationMutation rawPermutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":1,"startColumn":0,"endColumn":0},"sourceRows":[1,0]}
                """));
        OperationMutation permutation = withSortContext(rawPermutation, range(0, 1, 0, 0), "worksheet", null, false, 8);

        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation));

        assertEquals("SERVICE_UNAVAILABLE", error.code());
        assertEquals(before, snapshot);
    }

    @Test
    void rowPermutationRejectsFragmentedOutlineGroupWithoutChangingSnapshot() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":8,"columnCount":4,
                  "cells":{"0":{"0":{"value":"first"}},"1":{"0":{"value":"second"}}},"pane":{"kind":"none"},
                  "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},
                  "outline":{"groups":[{"id":"group-1","axis":"row","start":0,"end":1,"level":1,"collapsed":false}]},
                  "merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"sheetTables":[],"spillRanges":[],"protectionRules":[]}]}
                """);
        JsonNode before = snapshot.deepCopy();
        OperationMutation rawPermutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":3,"startColumn":0,"endColumn":1},"sourceRows":[2,0,3,1]}
                """));
        OperationMutation permutation = withSortContext(rawPermutation, range(0, 3, 0, 1), "worksheet", null, false, 3);

        ServiceException error = assertThrows(ServiceException.class,
                () -> registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation));

        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals(before, snapshot);
    }

    @Test
    void rowPermutationAllowsBoundedExactMetadataFragmentation() throws Exception {
        JsonNode current = applyFragmentedMetadataPermutation(256);

        assertEquals(256, current.path("sheets").get(0).path("conditionalFormats").get(0).path("ranges").size());
    }

    @Test
    void rowPermutationRejectsExcessiveExactMetadataFragmentation() throws Exception {
        ServiceException error = assertThrows(ServiceException.class, () -> applyFragmentedMetadataPermutation(257));

        assertEquals("VALIDATION_ERROR", error.code());
        assertEquals("Row permutation metadata produces too many exact ranges", error.getMessage());
    }

    private JsonNode applyFragmentedMetadataPermutation(int segmentCount) throws Exception {
        int rowCount = segmentCount * 2;
        ObjectNode snapshot = (ObjectNode) mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"Sheet1","columnCount":1,"cells":{},"review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"conditionalFormats":[{"id":"cf-1","sheetId":"sheet-1","ranges":[],"type":"highlight"}],"dataValidations":[],"sheetTables":[]}]}
                """);
        ObjectNode sheet = (ObjectNode) snapshot.path("sheets").get(0);
        sheet.put("rowCount", rowCount);
        ArrayNode ranges = (ArrayNode) sheet.path("conditionalFormats").get(0).path("ranges");
        ranges.add(mapper.readTree("{\"sheetId\":\"sheet-1\",\"startRow\":0,\"endRow\":" + (segmentCount - 1) + ",\"startColumn\":0,\"endColumn\":0}"));

        ArrayNode sourceRows = mapper.createArrayNode();
        int firstOutsideEvenRow = segmentCount + (segmentCount & 1);
        for (int row = 0; row < rowCount; row++) {
            if (row < segmentCount && (row & 1) == 1) sourceRows.add(firstOutsideEvenRow + row - 1);
            else if (row >= firstOutsideEvenRow && (row & 1) == 0) sourceRows.add(row - firstOutsideEvenRow + 1);
            else sourceRows.add(row);
        }
        ObjectNode params = mapper.createObjectNode();
        params.put("sheetId", "sheet-1");
        params.set("range", mapper.readTree("{\"sheetId\":\"sheet-1\",\"startRow\":0,\"endRow\":" + (rowCount - 1) + ",\"startColumn\":0,\"endColumn\":0}"));
        params.set("sourceRows", sourceRows);
        params.set("dataRegionContext", dataRegionContext(range(0, rowCount - 1, 0, 0), "worksheet", null, false));
        params.put("affectedColumnEnd", 0);
        OperationMutation permutation = new OperationMutation("rows.permuted", "sheet-1", params);

        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        return registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation);
    }

    @Test
    void tableDataBodySortUsesCanonicalFormulaOrderWithoutMovingHeaderOrTableOwner() throws Exception {
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        JsonNode snapshot = mapper.readTree("""
                {"sheets":[{"id":"sheet-1","name":"SortTableSheet","rowCount":6,"columnCount":3,
                  "cells":{"0":{"0":{"value":"Calculated"},"1":{"value":"Row"}},"1":{"0":{"value":20},"1":{"value":"twenty"}},"2":{"0":{"value":5},"1":{"value":"five"}},"3":{"0":{"value":10},"1":{"value":"ten"}}},
                  "pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                  "autoFilter":null,"sheetTables":[{"id":"table-1","sheetId":"sheet-1","name":"SortTable","range":{"sheetId":"sheet-1","startRow":0,"endRow":3,"startColumn":0,"endColumn":1},"hasHeaderRow":true,"hasTotalRow":false,"showBandedRows":false,"showBandedColumns":false,"showFirstColumn":false,"showLastColumn":false,"showFilterButton":true,"autoExpand":"both","autoFilter":{"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":0,"endRow":3,"startColumn":0,"endColumn":1},"columns":{}},"columns":[{"id":"calculated","name":"Calculated"},{"id":"row","name":"Row"}]}],
                  "review":{"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}},"hyperlinks":[],"merges":[],"conditionalFormats":[],"dataValidations":[],"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{},"spillRanges":[],"protectionRules":[]} ]}
                """);
        OperationMutation permutation = new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":1,"endRow":3,"startColumn":0,"endColumn":1},"sourceRows":[2,3,1]}
                """));
        permutation = withSortContext(permutation, range(0, 3, 0, 1), "sheet-table", "table-1", true, 2);

        JsonNode current = registry.prepare(snapshot, permutation, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, permutation);
        JsonNode sheet = current.path("sheets").get(0);
        assertEquals("Calculated", sheet.path("cells").path("0").path("0").path("value").asText());
        assertEquals("five", sheet.path("cells").path("1").path("1").path("value").asText());
        assertEquals("ten", sheet.path("cells").path("2").path("1").path("value").asText());
        assertEquals("twenty", sheet.path("cells").path("3").path("1").path("value").asText());
        assertEquals(0, sheet.path("sheetTables").get(0).path("range").path("startRow").asInt());
        assertEquals(3, sheet.path("sheetTables").get(0).path("range").path("endRow").asInt());
        assertEquals(0, sheet.path("sheetTables").get(0).path("autoFilter").path("range").path("startRow").asInt());

        OperationMutation duplicate = withSortContext(new OperationMutation("rows.permuted", "sheet-1", mapper.readTree("""
                {"sheetId":"sheet-1","range":{"sheetId":"sheet-1","startRow":1,"endRow":3,"startColumn":0,"endColumn":1},"sourceRows":[1,1,3]}
                """)), range(0, 3, 0, 1), "sheet-table", "table-1", true, 2);
        assertThrows(ServiceException.class, () -> registry.prepare(snapshot, duplicate, WorkbookAclRole.EDITOR).descriptor().apply(snapshot, duplicate));
    }

    @Test
    void sheetRenameRewritesAllPersistedFormulaOwnerCategories() throws Exception {
        ObjectNode snapshot = mapper.createObjectNode();
        ArrayNode sheets = snapshot.putArray("sheets");
        ObjectNode source = sheets.addObject().put("id", "source").put("name", "Source");
        ObjectNode sourceCells = source.putObject("cells");
        ObjectNode sourceRow = sourceCells.putObject("0");
        ObjectNode formulaCell = sourceRow.putObject("0");
        formulaCell.put("formula", "='Source'!A1");
        formulaCell.putObject("formulaMetadata").put("kind", "normal").put("sourceFormula", "='Source'!A1");
        ObjectNode barcodeCell = sourceRow.putObject("1");
        barcodeCell.putObject("presentation").put("kind", "barcode").putObject("source")
                .put("kind", "formula").put("formula", "='Source'!A1");
        ObjectNode tableSheet = source.putObject("tableSheet");
        tableSheet.putArray("columns").addObject().put("fieldId", "calculated").put("formula", "='Source'!A1");
        ObjectNode sourceDrawings = source.putObject("drawingPayloads");
        sourceDrawings.putObject("formula-shape").put("kind", "shape").put("propertyFormula", "='Source'!A1");
        sourceDrawings.putObject("chart").put("kind", "chart").put("chartId", "chart")
                .putObject("elements").put("hiddenData", "show").putObject("titleText").put("linkedFormula", "='Source'!A1");

        ObjectNode external = sheets.addObject().put("id", "owner").put("name", "Owner");
        external.putObject("cells").putObject("1").putObject("0").putObject("formulaMetadata")
                .put("kind", "normal").put("sourceFormula", "='Source'!A1");
        external.putArray("conditionalFormats").addObject().put("id", "cf").put("formula1", "='Source'!A1");
        external.putArray("dataValidations").addObject().put("id", "dv").put("listSource", mapper.createObjectNode()
                .put("kind", "formula").put("formula", "='Source'!A1"));
        ObjectNode externalDrawings = external.putObject("drawingPayloads");
        externalDrawings.putObject("external-shape").put("kind", "shape").put("propertyFormula", "='Source'!A1");
        externalDrawings.putObject("external-chart").put("kind", "chart").put("chartId", "external-chart")
                .putObject("elements").put("hiddenData", "show").putObject("legend").put("visible", true)
                .put("position", "bottom").putObject("text").put("linkedFormula", "='Source'!A1");

        snapshot.putArray("definedNameModels").addObject().put("name", "SourceName").put("formula", "='Source'!A1");
        snapshot.putObject("definedNames").put("SourceName", "='Source'!A1");
        snapshot.putObject("dataModel").putArray("views").addObject().put("id", "view")
                .putArray("fields").addObject().put("fieldId", "calculated").put("formula", "='Source'!A1");
        snapshot.putArray("cellStyleTemplates").addObject().put("id", "template").putObject("dataValidation")
                .put("formula1", "='Source'!A1");

        OperationMutation rename = new OperationMutation("sheet.rename", "source",
                mapper.readTree("{\"sheetId\":\"source\",\"name\":\"Renamed Sheet\"}"));
        JsonNode renamed = new WorkbookStructureMutationDescriptor("sheet.rename").apply(snapshot, rename);

        assertEquals("Renamed Sheet", renamed.path("sheets").get(0).path("name").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(0).path("cells").path("0").path("0").path("formula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(0).path("cells").path("0").path("0").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(0).path("cells").path("0").path("1").path("presentation").path("source").path("formula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(0).path("tableSheet").path("columns").get(0).path("formula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(0).path("drawingPayloads").path("formula-shape").path("propertyFormula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(0).path("drawingPayloads").path("chart").path("elements").path("titleText").path("linkedFormula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(1).path("cells").path("1").path("0").path("formulaMetadata").path("sourceFormula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(1).path("conditionalFormats").get(0).path("formula1").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(1).path("dataValidations").get(0).path("listSource").path("formula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(1).path("drawingPayloads").path("external-shape").path("propertyFormula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("sheets").get(1).path("drawingPayloads").path("external-chart").path("elements").path("legend").path("text").path("linkedFormula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("dataModel").path("views").get(0).path("fields").get(0).path("formula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("cellStyleTemplates").get(0).path("dataValidation").path("formula1").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("definedNameModels").get(0).path("formula").asText());
        assertEquals("='Renamed Sheet'!A1", renamed.path("definedNames").path("SourceName").asText());
        assertEquals("='Source'!A1", snapshot.path("sheets").get(0).path("cells").path("0").path("0").path("formulaMetadata").path("sourceFormula").asText());
    }

    @Test
    void sheetRenameRejectsPreservedOnlyFormulaOwnersWithoutChangingInput() throws Exception {
        ObjectNode snapshot = mapper.createObjectNode();
        ArrayNode sheets = snapshot.putArray("sheets");
        ObjectNode source = sheets.addObject().put("id", "source").put("name", "Source");
        source.putObject("cells").putObject("0").putObject("0").putObject("formulaMetadata")
                .put("kind", "dataTable").put("preservedOnly", true).put("sourceFormula", "='Source'!A1");
        OperationMutation rename = new OperationMutation("sheet.rename", "source",
                mapper.readTree("{\"sheetId\":\"source\",\"name\":\"Renamed\"}"));

        assertThrows(ServiceException.class, () -> new WorkbookStructureMutationDescriptor("sheet.rename").apply(snapshot, rename));
        assertEquals("Source", snapshot.path("sheets").get(0).path("name").asText());
        assertEquals("='Source'!A1", snapshot.path("sheets").get(0).path("cells").path("0").path("0").path("formulaMetadata").path("sourceFormula").asText());
    }

    @Test
    void sheetRenameRejectsMalformedChartLinkedFormulaWithoutChangingInput() throws Exception {
        ObjectNode snapshot = sheetDeletionSnapshot();
        ObjectNode chart = sheetDeletionOwner(snapshot).putObject("drawingPayloads").putObject("chart");
        chart.put("kind", "chart").put("chartId", "chart");
        chart.putObject("elements").put("hiddenData", "show").putObject("titleText").put("linkedFormula", 7);
        OperationMutation rename = new OperationMutation("sheet.rename", "source",
                mapper.readTree("{\"sheetId\":\"source\",\"name\":\"Renamed\"}"));

        assertThrows(ServiceException.class, () -> new WorkbookStructureMutationDescriptor("sheet.rename").apply(snapshot, rename));
        assertEquals("Source", snapshot.path("sheets").get(1).path("name").asText());
        assertEquals(7, snapshot.path("sheets").get(0).path("drawingPayloads").path("chart")
                .path("elements").path("titleText").path("linkedFormula").asInt());
    }

    @Test
    void sheetRemovalRejectsAuxiliaryFormulaAndWorkbookReferenceOwners() {
        ObjectNode formulaMetadata = sheetDeletionSnapshot();
        sheetDeletionOwner(formulaMetadata).putObject("cells").putObject("0").putObject("0").putObject("formulaMetadata")
                .put("sourceFormula", "='Source'!A1");

        ObjectNode barcode = sheetDeletionSnapshot();
        barcodeSourceCell(sheetDeletionOwner(barcode)).putObject("presentation").put("kind", "barcode")
                .putObject("source").put("kind", "formula").put("formula", "='Source'!A1");

        ObjectNode tableSheet = sheetDeletionSnapshot();
        sheetDeletionOwner(tableSheet).putObject("tableSheet").putArray("columns").addObject()
                .put("fieldId", "calculated").put("formula", "='Source'!A1");

        ObjectNode ruleFormula = sheetDeletionSnapshot();
        sheetDeletionOwner(ruleFormula).putArray("conditionalFormats").addObject().put("id", "cf").put("formula1", "='Source'!A1");

        ObjectNode viewFormula = sheetDeletionSnapshot();
        viewFormula.putObject("dataModel").putArray("views").addObject().put("id", "view")
                .putArray("fields").addObject().put("fieldId", "calculated").put("formula", "='Source'!A1");

        ObjectNode templateFormula = sheetDeletionSnapshot();
        templateFormula.putArray("cellStyleTemplates").addObject().put("id", "template").putObject("dataValidation")
                .put("formula1", "='Source'!A1");

        ObjectNode drawingFormula = sheetDeletionSnapshot();
        sheetDeletionOwner(drawingFormula).putObject("drawingPayloads").putObject("shape")
                .put("kind", "shape").put("propertyFormula", "='Source'!A1");

        ObjectNode tableRange = sheetDeletionSnapshot();
        sheetDeletionOwner(tableRange).putObject("autoFilter").putObject("range").put("sheetId", "source");

        ObjectNode queryTarget = sheetDeletionSnapshot();
        queryTarget.putArray("queryDefinitions").addObject().put("id", "query").putObject("lastTarget").put("sheetId", "source");

        ObjectNode printArea = sheetDeletionSnapshot();
        printArea.putArray("printDocuments").addObject().put("sheetId", "owner").putArray("printAreas")
                .addObject().put("sheetId", "source");

        ObjectNode definedNameFormula = sheetDeletionSnapshot();
        definedNameFormula.putArray("definedNameModels").addObject().put("name", "ExternalName")
                .put("scope", "workbook").put("formula", "='Source'!A1");

        ObjectNode definedNameAnchor = sheetDeletionSnapshot();
        definedNameAnchor.putArray("definedNameModels").addObject().put("name", "RelativeName")
                .put("scope", "workbook").put("formula", "=A1")
                .putObject("anchor").put("sheetId", "source").put("row", 0).put("column", 0);

        ObjectNode dataSource = sheetDeletionSnapshot();
        dataSource.putObject("dataModel").putArray("sources").addObject()
                .put("id", "source-data").put("sourceSheetId", "source");

        ObjectNode templateAnchor = sheetDeletionSnapshot();
        templateAnchor.putArray("cellStyleTemplates").addObject().put("id", "template")
                .putObject("dataValidation").putObject("formulaAnchor").put("sheetId", "source").put("row", 0).put("column", 0);

        ObjectNode reportTemplate = sheetDeletionSnapshot();
        sheetDeletionOwner(reportTemplate).putObject("reportSheet").put("templateSheetId", "source");

        ObjectNode shapeHyperlink = sheetDeletionSnapshot();
        shapeHyperlinkOwner(shapeHyperlink).put("kind", "shape").putObject("hyperlink")
                .put("kind", "sheet").put("sheetId", "source");

        ObjectNode chartSeriesRange = sheetDeletionSnapshot();
        ObjectNode chart = sheetDeletionOwner(chartSeriesRange).putObject("drawingPayloads").putObject("chart");
        chart.put("kind", "chart").putObject("source").put("kind", "worksheet-ranges").putArray("ranges")
                .addObject().put("sheetId", "owner");
        chart.putArray("series").addObject().putObject("stockRoles").putObject("high").put("sheetId", "source");

        ObjectNode chartFormula = sheetDeletionSnapshot();
        ObjectNode formulaChart = sheetDeletionOwner(chartFormula).putObject("drawingPayloads").putObject("chart");
        formulaChart.put("kind", "chart").putObject("source").put("kind", "worksheet-ranges").putArray("ranges");
        formulaChart.putObject("elements").put("hiddenData", "show").putObject("titleText")
                .put("linkedFormula", "='Source'!A1");

        for (ObjectNode snapshot : List.of(formulaMetadata, barcode, tableSheet, ruleFormula, viewFormula,
                templateFormula, drawingFormula, tableRange, queryTarget, printArea, definedNameFormula,
                definedNameAnchor, dataSource, templateAnchor, reportTemplate, shapeHyperlink, chartSeriesRange, chartFormula)) {
            assertSheetRemovalRejected(snapshot);
        }
    }

    @Test
    void sheetRemovalAllowsUnreferencedSheetAndRemovesItsScopedDocuments() throws Exception {
        ObjectNode snapshot = sheetDeletionSnapshot();
        snapshot.putArray("definedNameModels").addObject().put("name", "LocalName").put("scope", "sheet").put("sheetId", "source");
        snapshot.putArray("printDocuments").addObject().put("sheetId", "source");
        OperationMutation remove = new OperationMutation("sheet.remove", "source", mapper.createObjectNode().put("id", "source"));

        JsonNode next = new WorkbookStructureMutationDescriptor("sheet.remove").apply(snapshot, remove);

        assertEquals(1, next.path("sheets").size());
        assertEquals(0, next.path("definedNameModels").size());
        assertEquals(0, next.path("printDocuments").size());
        assertEquals(2, snapshot.path("sheets").size());
    }

    private void assertSheetRemovalRejected(ObjectNode snapshot) {
        OperationMutation remove = new OperationMutation("sheet.remove", "source", mapper.createObjectNode().put("id", "source"));
        assertThrows(ServiceException.class, () -> new WorkbookStructureMutationDescriptor("sheet.remove").apply(snapshot, remove));
        assertEquals(2, snapshot.path("sheets").size());
    }

    private ObjectNode sheetDeletionSnapshot() {
        ObjectNode root = mapper.createObjectNode();
        ArrayNode sheets = root.putArray("sheets");
        ObjectNode owner = sheets.addObject().put("id", "owner").put("name", "Owner");
        owner.putObject("cells");
        ObjectNode source = sheets.addObject().put("id", "source").put("name", "Source");
        source.putObject("cells");
        source.putArray("pivots");
        source.putArray("sheetTables");
        source.putArray("drawings");
        return root;
    }

    private ObjectNode sheetDeletionOwner(ObjectNode snapshot) {
        return (ObjectNode) snapshot.path("sheets").get(0);
    }

    private ObjectNode barcodeSourceCell(ObjectNode sheet) {
        ObjectNode cells = (ObjectNode) sheet.get("cells");
        ObjectNode row = (ObjectNode) cells.get("0");
        if (row == null) row = cells.putObject("0");
        return row.putObject("0");
    }

    private ObjectNode shapeHyperlinkOwner(ObjectNode snapshot) {
        return sheetDeletionOwner(snapshot).putObject("drawingPayloads").putObject("shape");
    }

    private ObjectNode drillDownDetail(String sheetId, String sourceId, int rowCount, List<String> headers) {
        ObjectNode detail = mapper.createObjectNode();
        ObjectNode source = detail.putObject("source");
        source.put("schema", "DataSourceManifest").put("version", 1).put("id", sourceId).put("name", "Details")
                .put("kind", "chunked-table").put("sourceSheetId", sheetId).put("rowCount", rowCount).put("blockRowCount", 65_536).put("revision", 0);
        source.putObject("sourceRange").put("sheetId", sheetId).put("startRow", 0).put("endRow", rowCount)
                .put("startColumn", 0).put("endColumn", headers.size() - 1);
        ArrayNode fields = source.putArray("fields");
        for (int index = 0; index < headers.size(); index++) fields.addObject().put("id", sourceId + ":field:" + index)
                .put("name", headers.get(index)).put("ordinal", index).put("type", "mixed");
        source.putArray("blocks").addObject().put("id", sourceId + ":block").put("dataSourceId", sourceId)
                .put("startRow", 0).put("rowCount", rowCount).put("storageKey", sourceId + ":block")
                .put("checksum", "b".repeat(64)).put("byteLength", 1).put("encoding", "columnar-v1").put("revision", 0);
        ObjectNode region = detail.putObject("region");
        region.put("id", sourceId + ":region").put("sourceId", sourceId).put("headerRow", 0).put("revision", 0);
        region.putObject("range").put("sheetId", sheetId).put("startRow", 0).put("endRow", rowCount)
                .put("startColumn", 0).put("endColumn", headers.size() - 1);
        ArrayNode detailHeaders = detail.putArray("headers");
        headers.forEach(detailHeaders::add);
        return detail;
    }

    private OperationMutation withDataRegionContext(OperationMutation mutation, ObjectNode range, String ownerKind, String tableId, boolean hasHeader) {
        ObjectNode params = (ObjectNode) mutation.params().deepCopy();
        params.set("dataRegionContext", dataRegionContext(range, ownerKind, tableId, hasHeader));
        return new OperationMutation(mutation.id(), mutation.sheetId(), params);
    }

    private OperationMutation withSortContext(OperationMutation mutation, ObjectNode range, String ownerKind, String tableId, boolean hasHeader, int affectedColumnEnd) {
        ObjectNode params = (ObjectNode) withDataRegionContext(mutation, range, ownerKind, tableId, hasHeader).params();
        params.put("affectedColumnEnd", affectedColumnEnd);
        return new OperationMutation(mutation.id(), mutation.sheetId(), params);
    }

    private ObjectNode dataRegionContext(ObjectNode range, String ownerKind, String tableId, boolean hasHeader) {
        ObjectNode context = mapper.createObjectNode().put("schema", "DataRegionContext").put("version", 1);
        context.set("selection", range.deepCopy());
        context.set("currentRegion", range.deepCopy());
        context.set("usedRange", range.deepCopy());
        context.set("range", range.deepCopy());
        ObjectNode owner = context.putObject("owner").put("kind", ownerKind);
        if (tableId != null) owner.put("tableId", tableId);
        ObjectNode header = context.putObject("header").put("kind", hasHeader ? "present" : "absent");
        if (hasHeader) header.put("row", range.path("startRow").asInt());
        context.putArray("visibleRows");
        return context;
    }

    private ObjectNode range(int startRow, int endRow, int startColumn, int endColumn) {
        return mapper.createObjectNode().put("sheetId", "sheet-1")
                .put("startRow", startRow).put("endRow", endRow)
                .put("startColumn", startColumn).put("endColumn", endColumn);
    }
}
