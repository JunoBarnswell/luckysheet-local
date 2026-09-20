package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.migration.SnapshotUpgrade;
import com.xc.luckysheet.server.mutation.MutationDescriptorRegistry;
import com.xc.luckysheet.server.store.CheckpointRow;
import com.xc.luckysheet.server.store.OperationRow;
import com.xc.luckysheet.server.store.WorkbookRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Duration;
import java.time.Instant;
import java.util.HexFormat;
import java.util.List;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyLong;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

/**
 * Service-level integrity and collaboration boundary checks.  The service is
 * wired with a real mutation registry; only the persistence boundary and ACL
 * are isolated so each assertion remains deterministic and transactional
 * decisions are exercised against canonical snapshots.
 */
class WorkbookIntegrityAndConflictTest {
    private static final String UNIT_ID = "book-1";
    private static final String SHEET_ID = "sheet-1";
    private static final String ACTOR = "local:actor-1";
    private static final Instant CREATED_AT = Instant.parse("2026-01-01T00:00:00Z");

    private final ObjectMapper mapper = new ObjectMapper().findAndRegisterModules();

    @Test
    void staleSameCellOperationIsRejectedByAffectedRangeConflict() throws Exception {
        String snapshot = canonicalSnapshot();
        WorkbookStore store = mock(WorkbookStore.class);
        AccessControlService access = mock(AccessControlService.class);
        OperationEnvelope previous = cellOperation("op-existing", "session-a", 1, 0, 0, 0, 1);
        OperationRow previousRow = committedRow(previous, 1, 0, 0);
        stubWorkbook(store, snapshot, 0, 1, List.of(previousRow));
        when(access.require(UNIT_ID, ACTOR, WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.EDITOR);

        WorkbookOperationService service = service(store, access, new MutationDescriptorRegistry());
        OperationEnvelope stale = cellOperation("op-stale", "session-b", 1, 0, 0, 0, 2);

        ServiceException error = assertThrows(ServiceException.class, () -> service.commit(UNIT_ID, stale, ACTOR));

        assertEquals("CONFLICT", error.code());
        assertTrue(error.getMessage().contains("Affected cells changed"));
        verify(store, never()).insertOperation(any());
    }

    @Test
    void staleNonIntersectingOperationCanCommitWhenItsMutationPolicyAllowsRebase() throws Exception {
        String snapshot = canonicalSnapshot();
        WorkbookStore store = mock(WorkbookStore.class);
        AccessControlService access = mock(AccessControlService.class);
        OperationEnvelope previous = cellOperation("op-existing", "session-a", 1, 0, 0, 0, 1);
        OperationRow previousRow = committedRow(previous, 1, 0, 0);
        stubWorkbook(store, snapshot, 0, 1, List.of(previousRow));
        when(access.require(UNIT_ID, ACTOR, WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.EDITOR);

        WorkbookOperationService service = service(store, access, new MutationDescriptorRegistry());
        OperationEnvelope nonIntersecting = cellOperation("op-next", "session-b", 1, 0, 0, 1, 2);

        WorkbookOperationService.CommitResult result = service.commit(UNIT_ID, nonIntersecting, ACTOR);

        assertTrue(result.committed());
        assertEquals(2, result.operation().revision());
        verify(store).insertOperation(any());
        verify(store).updateWorkbookRevisionAndName(eq(UNIT_ID), eq(2L), eq("Book"), any());
    }

    @Test
    void corruptedCheckpointFailsClosedInsteadOfReturningAnEmptyWorkbook() throws Exception {
        String snapshot = canonicalSnapshot();
        WorkbookStore store = mock(WorkbookStore.class);
        AccessControlService access = mock(AccessControlService.class);
        when(access.require(UNIT_ID, ACTOR, WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.VIEWER);
        when(store.find(UNIT_ID)).thenReturn(Optional.of(workbook(snapshot, 0, 0)));
        when(store.findCheckpoint(UNIT_ID, 0)).thenReturn(Optional.of(new CheckpointRow(
                UNIT_ID, 0, snapshot, "not-the-snapshot-checksum", CREATED_AT)));

        ServiceException error = assertThrows(ServiceException.class,
                () -> service(store, access, new MutationDescriptorRegistry()).readSnapshot(UNIT_ID, ACTOR));

        assertEquals("STORAGE_CORRUPT", error.code());
        assertEquals(409, error.status());
        assertTrue(error.getMessage().contains("Checksum mismatch"));
    }

    @Test
    void equalClientSequenceIsScopedToItsClientSession() throws Exception {
        String snapshot = canonicalSnapshot();
        WorkbookStore store = mock(WorkbookStore.class);
        AccessControlService access = mock(AccessControlService.class);
        stubWorkbook(store, snapshot, 0, 0, List.of());
        when(access.require(UNIT_ID, ACTOR, WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.EDITOR);
        when(store.findOperation(anyString())).thenReturn(Optional.empty());
        when(store.findOperationBySequence(anyString(), anyString(), anyString(), anyLong())).thenReturn(Optional.empty());

        WorkbookOperationService service = service(store, access, new MutationDescriptorRegistry());
        WorkbookOperationService.CommitResult first = service.commit(UNIT_ID,
                cellOperation("op-session-a", "session-a", 1, 0, 0, 0, 1), ACTOR);
        WorkbookOperationService.CommitResult second = service.commit(UNIT_ID,
                cellOperation("op-session-b", "session-b", 1, 0, 0, 1, 2), ACTOR);

        assertTrue(first.committed());
        assertTrue(second.committed());
        verify(store).findOperationBySequence(UNIT_ID, ACTOR, "session-a", 1);
        verify(store).findOperationBySequence(UNIT_ID, ACTOR, "session-b", 1);
    }

    @Test
    void duplicateClientSequenceInTheSameSessionIsRejected() throws Exception {
        String snapshot = canonicalSnapshot();
        WorkbookStore store = mock(WorkbookStore.class);
        AccessControlService access = mock(AccessControlService.class);
        OperationEnvelope existing = cellOperation("op-existing", "session-a", 1, 0, 0, 0, 1);
        stubWorkbook(store, snapshot, 0, 0, List.of());
        when(access.require(UNIT_ID, ACTOR, WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.EDITOR);
        when(store.findOperation("op-retry")).thenReturn(Optional.empty());
        when(store.findOperationBySequence(UNIT_ID, ACTOR, "session-a", 1))
                .thenReturn(Optional.of(committedRow(existing, 1, 0, 0)));

        WorkbookOperationService service = service(store, access, new MutationDescriptorRegistry());
        ServiceException error = assertThrows(ServiceException.class, () -> service.commit(UNIT_ID,
                cellOperation("op-retry", "session-a", 1, 0, 0, 1, 2), ACTOR));

        assertEquals("CONFLICT", error.code());
        assertTrue(error.getMessage().contains("clientSequence"));
        verify(store, never()).insertOperation(any());
    }

    private WorkbookOperationService service(WorkbookStore store, AccessControlService access,
                                             MutationDescriptorRegistry registry) {
        CoordinationProperties coordination = new CoordinationProperties(
                false, false, null, "coordination", Duration.ofSeconds(1), Duration.ofSeconds(30), 10, Duration.ofSeconds(45));
        return new WorkbookOperationService(store, access, registry, mapper,
                new AuditRecorder(store, mapper), coordination);
    }

    private void stubWorkbook(WorkbookStore store, String snapshot, long snapshotRevision, long revision,
                              List<OperationRow> operations) throws Exception {
        WorkbookRow row = workbook(snapshot, snapshotRevision, revision);
        when(store.findForUpdate(UNIT_ID)).thenReturn(Optional.of(row));
        when(store.findCheckpoint(UNIT_ID, snapshotRevision)).thenReturn(Optional.of(new CheckpointRow(
                UNIT_ID, snapshotRevision, snapshot, checksum(snapshot), CREATED_AT)));
        when(store.findOperation(anyString())).thenReturn(Optional.empty());
        when(store.findOperationBySequence(anyString(), anyString(), anyString(), anyLong())).thenReturn(Optional.empty());
        when(store.listOperations(UNIT_ID)).thenReturn(operations);
    }

    private WorkbookRow workbook(String snapshot, long snapshotRevision, long revision) {
        return new WorkbookRow(UNIT_ID, "Book", snapshot, snapshotRevision, revision,
                WorkbookLifecycle.ACTIVE, CREATED_AT, CREATED_AT);
    }

    private OperationEnvelope cellOperation(String operationId, String sessionId, long sequence, long baseRevision,
                                            int row, int column, int value) {
        ObjectNode params = mapper.createObjectNode()
                .put("sheetId", SHEET_ID)
                .put("row", row)
                .put("column", column);
        ObjectNode cellValue = params.putObject("value").put("value", value);
        ObjectNode authority = params.putObject("writeAuthority");
        authority.put("kind", "direct-entry");
        authority.putObject("target")
                .put("sheetId", SHEET_ID)
                .put("row", row)
                .put("column", column);
        authority.set("candidate", cellValue.deepCopy());
        authority.putObject("validationDecision").put("status", "accepted");
        return new OperationEnvelope(sessionId, OperationEnvelope.SCHEMA, operationId, UNIT_ID,
                sequence, baseRevision, List.of(new OperationMutation("cell.set", SHEET_ID, params)), CREATED_AT);
    }

    private OperationRow committedRow(OperationEnvelope operation, long revision, int row, int column) throws Exception {
        Instant committedAt = CREATED_AT.plusSeconds(revision);
        CommittedOperationMutation mutation = new CommittedOperationMutation(
                "cell.set", SHEET_ID, operation.mutations().get(0).params(),
                List.of(new RangeRef(SHEET_ID, row, row, column, column)));
        CommittedOperationEnvelope committed = CommittedOperationEnvelope.from(operation, ACTOR, revision,
                committedAt, List.of(mutation));
        return new OperationRow(operation.operationId(), UNIT_ID, revision, ACTOR,
                operation.clientSessionId(), operation.clientSequence(), operation.baseRevision(),
                mapper.writeValueAsString(committed), committedAt);
    }

    private String canonicalSnapshot() throws Exception {
        return mapper.writeValueAsString(SnapshotUpgrade.migrateStored(mapper.readTree("""
                {"schema":"WorkbookSnapshot","version":3,"unitId":"book-1","name":"Book",
                 "dimensionMetrics":{"normalFontFamily":"Calibri","normalFontSizePx":14.6666666667,"maximumDigitWidthPx":7},
                 "dataSources":[],"sheets":[{"id":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,
                 "cells":{},"merges":[],"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,
                 "pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{}}]}
                """), UNIT_ID));
    }

    private String checksum(String value) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256")
                .digest(value.getBytes(StandardCharsets.UTF_8)));
    }
}
