package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.service.ServiceException;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.mutation.MutationDescriptorRegistry;
import com.xc.luckysheet.server.store.WorkbookRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;

import java.time.Duration;
import java.time.Instant;
import java.util.List;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class WorkbookOperationServiceTest {
    private final ObjectMapper mapper = new ObjectMapper().findAndRegisterModules();

    @Test
    void inverseStructuralPatchSupportsUndoingACommittedCellShiftRestore() throws Exception {
        Instant committedAt = Instant.parse("2026-09-27T00:00:00Z");
        OperationMutation restore = new OperationMutation("cells.inserted.restore", "sheet-1", mapper.readTree("""
                {"spec":{"operation":"insert","axis":"row","range":{"sheetId":"sheet-1","startRow":2,"endRow":2,"startColumn":0,"endColumn":0}},"cells":{}}
                """));
        StructuralPatch patch = new StructuralPatch(StructuralPatch.VERSION, "cells.inserted.restore", List.of(), List.of(), List.of());
        OperationEnvelope request = new OperationEnvelope("test-session", OperationEnvelope.SCHEMA,
                "undo-op", "book-1", 1, 0, List.of(restore), committedAt);
        CommittedOperationEnvelope target = CommittedOperationEnvelope.from(request, "actor-1", 1, committedAt,
                List.of(CommittedOperationMutation.from(restore, List.of(), List.of(), patch)));
        OperationMutation redo = new OperationMutation("cells.inserted", "sheet-1", mapper.readTree("""
                {"operation":"insert","axis":"row","range":{"sheetId":"sheet-1","startRow":2,"endRow":2,"startColumn":0,"endColumn":0}}
                """));

        assertEquals(patch.inverse("cells.inserted"), WorkbookOperationService.inverseStructuralPatch(redo, target));
    }

    @Test
    void structuralUndoRejectsReusedInversesAndMutationsOutsideAnAllStructuralTarget() throws Exception {
        Instant committedAt = Instant.parse("2026-09-27T00:00:00Z");
        OperationMutation insertion = new OperationMutation("rows.inserted", "sheet-1",
                mapper.readTree("""{"at":2,"count":1}"""));
        StructuralPatch patch = new StructuralPatch(StructuralPatch.VERSION, "rows.inserted", List.of(), List.of(), List.of());
        OperationEnvelope targetRequest = new OperationEnvelope("test-session", OperationEnvelope.SCHEMA,
                "structural-target", "book-1", 1, 0, List.of(insertion), committedAt);
        CommittedOperationMutation committedInsertion = CommittedOperationMutation.from(insertion, List.of(), List.of(), patch);
        CommittedOperationEnvelope target = CommittedOperationEnvelope.from(targetRequest, "actor-1", 1, committedAt,
                List.of(committedInsertion));
        OperationMutation deletion = new OperationMutation("rows.deleted", "sheet-1",
                mapper.readTree("""{"at":2,"count":1}"""));
        OperationMutation unrelated = new OperationMutation("cell.set", "sheet-1",
                mapper.readTree("""{"row":0,"column":0,"value":"extra"}"""));
        OperationEnvelope withExtraMutation = new OperationEnvelope("test-session", OperationEnvelope.SCHEMA,
                "structural-undo-extra", "book-1", 2, 1, List.of(deletion, unrelated), committedAt);

        assertThrows(ServiceException.class,
                () -> WorkbookOperationService.validateStructuralUndoMutations(withExtraMutation, target, null));

        OperationEnvelope duplicateTargetRequest = new OperationEnvelope("test-session", OperationEnvelope.SCHEMA,
                "duplicate-structural-target", "book-1", 1, 0, List.of(insertion, insertion), committedAt);
        CommittedOperationEnvelope duplicateTarget = CommittedOperationEnvelope.from(duplicateTargetRequest, "actor-1", 1,
                committedAt, List.of(committedInsertion, committedInsertion));
        OperationEnvelope reusedInverse = new OperationEnvelope("test-session", OperationEnvelope.SCHEMA,
                "structural-undo-reused", "book-1", 2, 1, List.of(deletion), committedAt);

        assertThrows(ServiceException.class,
                () -> WorkbookOperationService.validateStructuralUndoMutations(reusedInverse, duplicateTarget, null));
    }

    @Test
    void commenterCannotCommitAnEditorMutationEvenThoughTheRequestHasNoClientRole() throws Exception {
        WorkbookStore store = mock(WorkbookStore.class);
        AccessControlService access = mock(AccessControlService.class);
        AuditRecorder audit = mock(AuditRecorder.class);
        CoordinationProperties coordination = new CoordinationProperties(
                false, false, null, "coordination", Duration.ofSeconds(1), Duration.ofSeconds(30), 10, Duration.ofSeconds(45)
        );
        WorkbookOperationService service = new WorkbookOperationService(
                store, access, new MutationDescriptorRegistry(), mapper, audit, coordination,
                mock(WorkbookDataBlockPublicationGuard.class)
        );
        String snapshot = mapper.writeValueAsString(com.xc.luckysheet.server.migration.SnapshotUpgrade.migrateStored(mapper.readTree(canonicalSnapshot()), "book-1"));
        when(store.findCheckpoint("book-1", 0)).thenReturn(Optional.of(new com.xc.luckysheet.server.store.CheckpointRow("book-1", 0, snapshot,
                java.util.HexFormat.of().formatHex(java.security.MessageDigest.getInstance("SHA-256").digest(snapshot.getBytes(java.nio.charset.StandardCharsets.UTF_8))), Instant.now())));
        when(access.require("book-1", "guest:share-1", WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.COMMENTER);
        when(store.findForUpdate("book-1")).thenReturn(Optional.of(new WorkbookRow(
                "book-1", "Book", snapshot, 0, 0, WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now()
        )));
        when(store.findOperation("op-1")).thenReturn(Optional.empty());
        when(store.findOperationBySequence("book-1", "guest:share-1", "test-session", 1)).thenReturn(Optional.empty());

        OperationEnvelope operation = new OperationEnvelope("test-session", 
                OperationEnvelope.SCHEMA,
                "op-1",
                "book-1",
                1,
                0,
                List.of(new OperationMutation("cell.set", "sheet-1", mapper.readTree("{\"row\":0,\"column\":0,\"value\":{\"value\":1}}"))),
                Instant.parse("2000-01-01T00:00:00Z")
        );

        ServiceException error = assertThrows(ServiceException.class, () -> service.commit("book-1", operation, "guest:share-1"));
        assertEquals("FORBIDDEN", error.code());
        verify(store, never()).insertOperation(any());
        verify(audit).rejected(eq("op-1"), eq("book-1"), eq("guest:share-1"), eq("OPERATION_COMMIT"), any());
    }

    @Test
    void serverAddsActorTimeAndRangesOnlyAfterItHasValidatedTheMutation() throws Exception {
        WorkbookStore store = mock(WorkbookStore.class);
        AccessControlService access = mock(AccessControlService.class);
        AuditRecorder audit = mock(AuditRecorder.class);
        WorkbookOperationService service = new WorkbookOperationService(
                store,
                access,
                new MutationDescriptorRegistry(),
                mapper,
                new AuditRecorder(store, mapper),
                new CoordinationProperties(false, false, null, "coordination", Duration.ofSeconds(1), Duration.ofSeconds(30), 10, Duration.ofSeconds(45)),
                mock(WorkbookDataBlockPublicationGuard.class)
        );
        String snapshot = mapper.writeValueAsString(com.xc.luckysheet.server.migration.SnapshotUpgrade.migrateStored(mapper.readTree(canonicalSnapshot()), "book-1"));
        when(store.findCheckpoint("book-1", 0)).thenReturn(Optional.of(new com.xc.luckysheet.server.store.CheckpointRow("book-1", 0, snapshot,
                java.util.HexFormat.of().formatHex(java.security.MessageDigest.getInstance("SHA-256").digest(snapshot.getBytes(java.nio.charset.StandardCharsets.UTF_8))), Instant.now())));
        when(access.require("book-1", "editor-1", WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.EDITOR);
        when(store.findForUpdate("book-1")).thenReturn(Optional.of(new WorkbookRow("book-1", "Book", snapshot, 0, 0,
                WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now())));
        when(store.findOperation("op-2")).thenReturn(Optional.empty());
        when(store.findOperationBySequence("book-1", "editor-1", "test-session", 1)).thenReturn(Optional.empty());
        when(store.listOperations("book-1")).thenReturn(List.of());
        OperationEnvelope operation = new OperationEnvelope("test-session", 
                OperationEnvelope.SCHEMA,
                "op-2",
                "book-1",
                1,
                0,
                List.of(new OperationMutation("cell.set", "sheet-1", mapper.readTree("{\"sheetId\":\"sheet-1\",\"row\":1,\"column\":2,\"value\":{\"value\":42},\"writeAuthority\":{\"kind\":\"script\",\"target\":{\"sheetId\":\"sheet-1\",\"row\":1,\"column\":2},\"candidate\":{\"value\":42},\"validationDecision\":{\"status\":\"accepted\"}}}"))),
                Instant.parse("2000-01-01T00:00:00Z")
        );

        WorkbookOperationService.CommitResult result = service.commit("book-1", operation, "editor-1");

        assertEquals(true, result.committed());
        assertEquals("editor-1", result.operation().actorId());
        assertEquals(result.operation().committedAt(), result.operation().createdAt());
        assertEquals(1, result.operation().mutations().get(0).affectedRanges().get(0).startRow());
        assertEquals(2, result.operation().mutations().get(0).affectedRanges().get(0).startColumn());
        ArgumentCaptor<com.xc.luckysheet.server.store.OperationRow> captured = ArgumentCaptor.forClass(com.xc.luckysheet.server.store.OperationRow.class);
        verify(store).insertOperation(captured.capture());
        assertEquals("op-2", captured.getValue().operationId());
        verify(store).updateWorkbookRevisionAndName(eq("book-1"), eq(1L), eq("Book"), any());
    }

    @Test
    void rowPermutationUndoRequiresTheExactInverseSourceOrder() throws Exception {
        var original = mapper.readTree("""
                {"range":{"sheetId":"sheet-1","startRow":5,"endRow":7,"startColumn":0,"endColumn":2},
                 "affectedColumnEnd":9,"sourceRows":[7,5,6]}
                """);
        var inverse = mapper.readTree("""
                {"range":{"sheetId":"sheet-1","startRow":5,"endRow":7,"startColumn":0,"endColumn":2},
                 "affectedColumnEnd":9,"sourceRows":[6,7,5]}
                """);

        assertTrue(WorkbookOperationService.sameRowPermutationInverse(original, inverse));
        assertFalse(WorkbookOperationService.sameRowPermutationInverse(original, mapper.readTree("""
                {"range":{"sheetId":"sheet-1","startRow":5,"endRow":7,"startColumn":0,"endColumn":2},
                 "affectedColumnEnd":9,"sourceRows":[7,5,6]}
                """)));
        assertFalse(WorkbookOperationService.sameRowPermutationInverse(original, mapper.readTree("""
                {"range":{"sheetId":"sheet-1","startRow":5,"endRow":7,"startColumn":0,"endColumn":2},
                 "affectedColumnEnd":8,"sourceRows":[6,7,5]}
                """)));
    }

    private String canonicalSnapshot() {
        return "{\"schema\":\"WorkbookSnapshot\",\"version\":3,\"unitId\":\"book-1\",\"name\":\"Book\",\"dimensionMetrics\":{\"normalFontFamily\":\"Calibri\",\"normalFontSizePx\":14.6666666667,\"maximumDigitWidthPx\":7},\"dataSources\":[],\"sheets\":[{\"id\":\"sheet-1\",\"name\":\"Sheet1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{},\"merges\":[],\"pane\":{\"kind\":\"none\"},\"defaultRowHeightPx\":20,\"defaultColumnWidthPx\":64,\"pivots\":[],\"sparklines\":[],\"drawings\":[],\"drawingPayloads\":{}}]}";
    }
}
