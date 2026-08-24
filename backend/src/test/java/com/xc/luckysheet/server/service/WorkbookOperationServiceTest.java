package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationMutation;
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
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class WorkbookOperationServiceTest {
    private final ObjectMapper mapper = new ObjectMapper().findAndRegisterModules();

    @Test
    void commenterCannotCommitAnEditorMutationEvenThoughTheRequestHasNoClientRole() throws Exception {
        WorkbookStore store = mock(WorkbookStore.class);
        AccessControlService access = mock(AccessControlService.class);
        AuditRecorder audit = mock(AuditRecorder.class);
        CoordinationProperties coordination = new CoordinationProperties(
                false, false, null, "coordination", Duration.ofSeconds(1), Duration.ofSeconds(30), 10, Duration.ofSeconds(45)
        );
        WorkbookOperationService service = new WorkbookOperationService(
                store, access, new MutationDescriptorRegistry(), mapper, audit, coordination
        );
        String snapshot = canonicalSnapshot();
        when(access.require("book-1", "guest:share-1", WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.COMMENTER);
        when(store.findForUpdate("book-1")).thenReturn(Optional.of(new WorkbookRow(
                "book-1", "Book", snapshot, 0, 0, WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now()
        )));
        when(store.findOperation("op-1")).thenReturn(Optional.empty());
        when(store.findOperationBySequence("book-1", "guest:share-1", 1)).thenReturn(Optional.empty());

        OperationEnvelope operation = new OperationEnvelope(
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
                new CoordinationProperties(false, false, null, "coordination", Duration.ofSeconds(1), Duration.ofSeconds(30), 10, Duration.ofSeconds(45))
        );
        String snapshot = canonicalSnapshot();
        when(access.require("book-1", "editor-1", WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.EDITOR);
        when(store.findForUpdate("book-1")).thenReturn(Optional.of(new WorkbookRow("book-1", "Book", snapshot, 0, 0,
                WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now())));
        when(store.findOperation("op-2")).thenReturn(Optional.empty());
        when(store.findOperationBySequence("book-1", "editor-1", 1)).thenReturn(Optional.empty());
        when(store.listOperations("book-1")).thenReturn(List.of());
        OperationEnvelope operation = new OperationEnvelope(
                OperationEnvelope.SCHEMA,
                "op-2",
                "book-1",
                1,
                0,
                List.of(new OperationMutation("cell.set", "sheet-1", mapper.readTree("{\"row\":1,\"column\":2,\"value\":{\"value\":42}}"))),
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

    private String canonicalSnapshot() {
        return "{\"schema\":\"WorkbookSnapshot\",\"version\":3,\"unitId\":\"book-1\",\"name\":\"Book\",\"dimensionMetrics\":{\"normalFontFamily\":\"Calibri\",\"normalFontSizePx\":14.6666666667,\"maximumDigitWidthPx\":7},\"dataSources\":[],\"sheets\":[{\"id\":\"sheet-1\",\"name\":\"Sheet1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{},\"merges\":[],\"pane\":{\"kind\":\"none\"},\"defaultRowHeightPx\":20,\"defaultColumnWidthPx\":64,\"pivots\":[],\"sparklines\":[],\"drawings\":[],\"drawingPayloads\":{}}]}";
    }
}
