package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.AclEntry;
import com.xc.luckysheet.server.contract.AuditRecord;
import com.xc.luckysheet.server.contract.CheckpointResponse;
import com.xc.luckysheet.server.contract.CursorPage;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationIntent;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.RestoreRequest;
import com.xc.luckysheet.server.contract.RevisionRecord;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookAccessProjection;
import com.xc.luckysheet.server.contract.WorkbookSnapshotResponse;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.contract.WorkbookSnapshotValidator;
import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.mutation.MutationDescriptorRegistry;
import com.xc.luckysheet.server.mutation.MutationPreparation;
import com.xc.luckysheet.server.store.CheckpointRow;
import com.xc.luckysheet.server.store.OperationRow;
import com.xc.luckysheet.server.store.OutboxRow;
import com.xc.luckysheet.server.store.WorkbookRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;

@Service
public class WorkbookOperationService {
    private static final long CHECKPOINT_OPERATION_LIMIT = 50;
    private static final long CHECKPOINT_BYTES_LIMIT = 512_000;
    private static final String SYSTEM_RESTORE_ACTOR = "system:workbook-restore";

    private final WorkbookStore store;
    private final AccessControlService access;
    private final MutationDescriptorRegistry registry;
    private final ObjectMapper mapper;
    private final AuditRecorder auditRecorder;
    private final CoordinationProperties coordination;
    /**
     * H2 runs the browser service as one JVM.  Keep commit, checkpoint and
     * restore writes for one workbook in one local critical section so a
     * checkpoint cannot race the operation replay entity update.  The
     * database lock remains authoritative for future instances.
     */
    private final ConcurrentHashMap<String, Object> workbookLocks = new ConcurrentHashMap<>();

    public WorkbookOperationService(
            WorkbookStore store,
            AccessControlService access,
            MutationDescriptorRegistry registry,
            ObjectMapper mapper,
            AuditRecorder auditRecorder,
            CoordinationProperties coordination
    ) {
        this.store = store;
        this.access = access;
        this.registry = registry;
        this.mapper = mapper;
        this.auditRecorder = auditRecorder;
        this.coordination = coordination;
    }

    public CommitResult operationResult(String unitId, String operationId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        OperationRow row = store.findOperation(operationId).orElseThrow(() -> ServiceException.notFound("Operation not committed"));
        if (!row.unitId().equals(unitId) || !row.actorSubject().equals(actor)) throw ServiceException.forbidden("Operation belongs to another subject");
        return new CommitResult(readCommitted(row), false);
    }

    public WorkbookSnapshotResponse readSnapshot(String unitId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookRow row = requireWorkbook(unitId);
        JsonNode snapshot = currentSnapshot(row);
        if (snapshot.isObject()) ((ObjectNode) snapshot).put("name", row.name());
        String json = writeJson(snapshot);
        return response(unitId, snapshot, row.revision(), checksum(json));
    }

    @Transactional
    public CommitResult commit(String routeUnitId, OperationEnvelope operation, String actor) {
        synchronized (workbookLock(routeUnitId)) {
            try {
                if (operation == null) throw ServiceException.validation("Operation is required");
                return commitInternal(routeUnitId, operation, actor);
            } catch (ServiceException error) {
                auditRecorder.rejected(operation == null ? null : operation.operationId(), routeUnitId, actor, "OPERATION_COMMIT", error.getMessage());
                throw error;
            }
        }
    }

    private CommitResult commitInternal(String routeUnitId, OperationEnvelope operation, String actor) {
        if (!routeUnitId.equals(operation.unitId())) throw ServiceException.validation("Operation unitId does not match route");
        WorkbookAclRole actorRole = access.require(routeUnitId, actor, WorkbookAclRole.VIEWER);
        WorkbookRow row = store.findForUpdate(routeUnitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + routeUnitId));
        if (row.lifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash and cannot accept operations");

        OperationRow existing = store.findOperation(operation.operationId()).orElse(null);
        if (existing != null) {
            if (!existing.actorSubject().equals(actor) || !existing.unitId().equals(routeUnitId)) {
                throw ServiceException.forbidden("Operation belongs to another subject");
            }
            CommittedOperationEnvelope committed = readCommitted(existing);
            List<OperationMutation> originalMutations = committed.mutations().stream()
                    .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList();
            // Server-owned timestamps/ranges are deliberately excluded from request identity.
            if (!committed.clientSessionId().equals(operation.clientSessionId())
                    || committed.clientSequence() != operation.clientSequence()
                    || committed.baseRevision() != operation.baseRevision()
                    || !originalMutations.equals(operation.mutations())
                    || !java.util.Objects.equals(committed.intent(), operation.intent())) {
                throw new ServiceException("OPERATION_ID_REUSED", 409,
                        "Operation " + operation.operationId() + " was already committed with different content; retain the draft and create a new operation");
            }
            return new CommitResult(committed, false);
        }
        OperationRow sequenceExisting = store.findOperationBySequence(routeUnitId, actor, operation.clientSessionId(), operation.clientSequence()).orElse(null);
        if (sequenceExisting != null && !sequenceExisting.operationId().equals(operation.operationId())) {
            throw ServiceException.conflict("clientSequence was already committed");
        }
        validateIntent(routeUnitId, operation, actor, row);
        if (operation.baseRevision() > row.revision()) {
            throw ServiceException.conflict("Base revision is ahead of the server; reload before submitting");
        }
        if (registry.requiresExactBase(operation.mutations()) && row.revision() != operation.baseRevision()) {
            throw ServiceException.conflict("Revision conflict; rebase against revision " + row.revision());
        }

        JsonNode before = currentSnapshot(row);
        JsonNode next = before;
        List<CommittedOperationMutation> committedMutations = new ArrayList<>();
        for (OperationMutation mutation : operation.mutations()) {
            MutationPreparation prepared = registry.prepare(next, mutation, actorRole);
            next = prepared.descriptor().apply(next, mutation);
            committedMutations.add(CommittedOperationMutation.from(mutation, prepared.affectedRanges()));
        }

        if (operation.baseRevision() < row.revision()) {
            for (OperationRow intervening : store.listOperationsBetween(routeUnitId, operation.baseRevision(), row.revision())) {
                var interveningMutations = readCommitted(intervening).mutations();
                if (registry.requiresExactBase(interveningMutations.stream()
                        .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList())) {
                    throw ServiceException.conflict("Workbook structure changed after base revision; reload and review the draft");
                }
                for (var previous : interveningMutations) {
                    for (var proposed : committedMutations) {
                        for (var left : previous.affectedRanges()) for (var right : proposed.affectedRanges()) {
                            if (left.sheetId().equals(right.sheetId()) && left.startRow() <= right.endRow() && left.endRow() >= right.startRow()
                                    && left.startColumn() <= right.endColumn() && left.endColumn() >= right.startColumn()) {
                                throw ServiceException.conflict("Affected cells changed after base revision; reload and review the draft");
                            }
                        }
                    }
                }
            }
        }

        long nextRevision = row.revision() + 1;
        Instant committedAt = Instant.now();
        CommittedOperationEnvelope committed = CommittedOperationEnvelope.from(operation, actor, nextRevision, committedAt, committedMutations);
        String envelopeJson = writeJson(committed);
        store.insertOperation(new OperationRow(operation.operationId(), routeUnitId, nextRevision, actor, operation.clientSessionId(), operation.clientSequence(), operation.baseRevision(), envelopeJson, committedAt));
        enqueueRevisionEvent(routeUnitId, operation.operationId(), nextRevision, envelopeJson, committedAt);
        String canonicalName = next.path("name").asText(row.name()).trim();
        store.updateWorkbookRevisionAndName(routeUnitId, nextRevision, canonicalName, committedAt);
        if (shouldCheckpoint(row, operation, envelopeJson)) {
            String nextJson = writeJson(next);
            store.updateWorkbook(routeUnitId, nextRevision, nextJson, nextRevision, committedAt);
            store.insertCheckpoint(routeUnitId, nextRevision, nextJson, checksum(nextJson), committedAt);
        }
        audit(operation.operationId(), routeUnitId, actor, "OPERATION_COMMIT", "ACCEPTED", null, mapper.createObjectNode().put("revision", nextRevision));
        return new CommitResult(committed, true);
    }

    private void validateIntent(String unitId, OperationEnvelope operation, String actor, WorkbookRow row) {
        OperationIntent intent = operation.intent();
        if (intent == null) return;
        if (!OperationIntent.UNDO.equals(intent.type())) {
            throw ServiceException.validation("Unsupported operation intent");
        }
        if (intent.targetOperationId().equals(operation.operationId())) {
            throw ServiceException.conflict("Undo operation cannot target itself");
        }
        OperationRow target = store.findOperation(intent.targetOperationId()).orElseThrow(
                () -> ServiceException.conflict("Undo target operation is not committed"));
        if (!unitId.equals(target.unitId())) {
            throw ServiceException.forbidden("Undo target belongs to another workbook");
        }
        if (!actor.equals(target.actorSubject())) {
            throw ServiceException.forbidden("Undo target belongs to another subject");
        }
        if (intent.targetBaseRevision() != target.baseRevision()) {
            throw ServiceException.conflict("Undo target base revision does not match the committed operation");
        }
        if (operation.baseRevision() != row.revision()) {
            throw ServiceException.conflict("Undo requires the current workbook revision " + row.revision());
        }
    }

    public CursorPage<RevisionRecord> revisions(String unitId, String actor, long beforeRevision, int limit, String nextCursor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        List<RevisionRecord> items = store.listOperationsBefore(unitId, beforeRevision, limit).stream().map(this::revisionRecord).toList();
        String next = items.size() == limit ? Long.toString(items.get(items.size() - 1).revision()) : null;
        return new CursorPage<>(items, next);
    }

    public WorkbookSnapshotResponse readRevision(String unitId, long revision, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookRow current = requireWorkbook(unitId);
        if (revision < 0 || revision > current.revision()) throw ServiceException.notFound("Revision not found: " + revision);
        JsonNode snapshot = snapshotAtRevision(current, revision);
        String json = writeJson(snapshot);
        return response(unitId, snapshot, revision, checksum(json));
    }

    @Transactional
    public CheckpointResponse checkpoint(String unitId, String actor) {
        synchronized (workbookLock(unitId)) {
            access.require(unitId, actor, WorkbookAclRole.EDITOR);
            WorkbookRow row = store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
            if (row.lifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash and cannot be checkpointed");
            if (row.snapshotRevision() == row.revision()) {
                JsonNode snapshot = currentSnapshot(row);
                String canonicalJson = writeJson(snapshot);
                return new CheckpointResponse(unitId, row.revision(), checksum(canonicalJson), false);
            }
            JsonNode snapshot = currentSnapshot(row);
            String json = writeJson(snapshot);
            Instant now = Instant.now();
            store.updateWorkbook(unitId, row.revision(), json, row.revision(), now);
            store.insertCheckpoint(unitId, row.revision(), json, checksum(json), now);
            return new CheckpointResponse(unitId, row.revision(), checksum(json), true);
        }
    }

    @Transactional
    public RestoreResult restore(String unitId, RestoreRequest request, String actor) {
        synchronized (workbookLock(unitId)) {
            access.require(unitId, actor, WorkbookAclRole.OWNER);
            WorkbookRow row = store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
            if (request.targetRevision() > row.revision()) throw ServiceException.notFound("Revision not found: " + request.targetRevision());
            JsonNode target = snapshotAtRevision(row, request.targetRevision());
            WorkbookSnapshotValidator.requireCanonical(target, unitId);
            registry.require("workbook.restore", true);
            long revision = row.revision() + 1;
            Instant now = Instant.now();
            String operationId = UUID.randomUUID().toString();
            ObjectNode params = mapper.createObjectNode()
                    .put("serverGenerated", true)
                    .put("targetRevision", request.targetRevision())
                    .put("reason", request.reason());
            params.set("snapshot", target.deepCopy());
            List<CommittedOperationMutation> mutations = fullWorkbookRestoreMutation(target, params);
            List<OperationMutation> sourceMutations = mutations.stream()
                    .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList();
            OperationEnvelope source = new OperationEnvelope("system", OperationEnvelope.SCHEMA, operationId, unitId, revision, row.revision(),
                    sourceMutations, now);
            CommittedOperationEnvelope committed = CommittedOperationEnvelope.system(source, actor, revision, now, mutations);
            String json = writeJson(target);
            String envelopeJson = writeJson(committed);
            // System restores must not share an authenticated browser's client-sequence namespace.
            store.insertOperation(new OperationRow(operationId, unitId, revision, SYSTEM_RESTORE_ACTOR, "system", revision, row.revision(), envelopeJson, now));
            enqueueRevisionEvent(unitId, operationId, revision, envelopeJson, now);
            store.updateWorkbook(unitId, revision, json, revision, now);
            store.insertCheckpoint(unitId, revision, json, checksum(json), now);
            audit(operationId, unitId, actor, "SNAPSHOT_RESTORE", "ACCEPTED", request.reason(), mapper.createObjectNode().put("targetRevision", request.targetRevision()));
            return new RestoreResult(committed, response(unitId, target, revision, checksum(json)));
        }
    }

    private Object workbookLock(String unitId) {
        if (unitId == null || unitId.isBlank()) return this;
        return workbookLocks.computeIfAbsent(unitId, ignored -> new Object());
    }

    public List<AclEntry> acl(String unitId, String actor) {
        return access.list(unitId, actor);
    }

    /** Read-only projection; it never accepts a browser-declared actor or role. */
    public WorkbookAccessProjection accessProjection(String unitId, String actor) {
        return new WorkbookAccessProjection(unitId, access.currentRole(unitId, actor));
    }

    public AclEntry grantAcl(String unitId, String actor, String target, WorkbookAclRole role) {
        return access.grant(unitId, actor, target, role);
    }

    public void revokeAcl(String unitId, String actor, String target) {
        access.revoke(unitId, actor, target);
    }

    public List<AuditRecord> audit(String unitId, String actor, int limit) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        requireWorkbook(unitId);
        return store.listAudit(unitId, limit);
    }

    private List<CommittedOperationMutation> fullWorkbookRestoreMutation(JsonNode target, ObjectNode params) {
        List<CommittedOperationMutation> mutations = new ArrayList<>();
        JsonNode sheets = target.path("sheets");
        if (sheets.isArray()) {
            for (JsonNode sheet : sheets) {
                String sheetId = sheet.path("id").asText("unknown");
                mutations.add(new CommittedOperationMutation("workbook.restore", sheetId, params, List.of(new RangeRef(sheetId, 0, 1_048_575, 0, 16_383))));
            }
        }
        if (mutations.isEmpty()) mutations.add(new CommittedOperationMutation("workbook.restore", "workbook", params, List.of()));
        return mutations;
    }

    private RevisionRecord revisionRecord(OperationRow row) {
        CommittedOperationEnvelope operation = readCommitted(row);
        return new RevisionRecord(row.operationId(), row.revision(), row.committedAt(), operation);
    }

    private CommittedOperationEnvelope readCommitted(OperationRow row) {
        try {
            return mapper.readValue(row.envelopeJson(), CommittedOperationEnvelope.class);
        } catch (Exception error) {
            throw new IllegalStateException("Stored operation envelope is invalid", error);
        }
    }

    private JsonNode currentSnapshot(WorkbookRow row) {
        CheckpointRow checkpoint = store.findCheckpoint(row.unitId(), row.snapshotRevision())
                .orElseThrow(() -> new ServiceException("STORAGE_CORRUPT", 409, "Checkpoint missing for " + row.unitId() + "; restore a verified backup"));
        verifyCheckpoint(checkpoint);
        if (!checksum(row.snapshotJson()).equals(checkpoint.checksum())) {
            throw new ServiceException("STORAGE_CORRUPT", 409, "Snapshot/checkpoint mismatch for " + row.unitId() + "; restore a verified backup");
        }
        JsonNode snapshot = WorkbookSnapshotValidator.requireCanonical(readJson(row.snapshotJson()), row.unitId());
        if (row.snapshotRevision() == row.revision()) return snapshot;
        for (OperationRow operation : store.listOperationsBetween(row.unitId(), row.snapshotRevision(), row.revision())) {
            CommittedOperationEnvelope committed = readCommitted(operation);
            List<OperationMutation> mutations = committed.mutations().stream()
                    .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList();
            snapshot = registry.applyPublicMutations(snapshot, mutations);
        }
        return snapshot;
    }

    private JsonNode snapshotAtRevision(WorkbookRow current, long targetRevision) {
        if (targetRevision == current.revision()) return currentSnapshot(current);
        CheckpointRow checkpoint = store.findCheckpoint(current.unitId(), targetRevision)
                .orElseGet(() -> store.findLatestCheckpointAtOrBefore(current.unitId(), targetRevision).orElseThrow(() -> ServiceException.notFound("Snapshot checkpoint not found")));
        verifyCheckpoint(checkpoint);
        JsonNode snapshot = WorkbookSnapshotValidator.requireCanonical(readJson(checkpoint.snapshotJson()), current.unitId());
        if (checkpoint.revision() == targetRevision) return snapshot;
        for (OperationRow operation : store.listOperationsBetween(current.unitId(), checkpoint.revision(), targetRevision)) {
            CommittedOperationEnvelope committed = readCommitted(operation);
            if (committed.mutations().stream().anyMatch(mutation -> "workbook.restore".equals(mutation.id()))) {
                throw ServiceException.conflict("Restore checkpoint is missing for revision " + operation.revision());
            }
            snapshot = registry.applyPublicMutations(snapshot, committed.mutations().stream()
                    .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList());
        }
        return snapshot;
    }

    private WorkbookRow requireWorkbook(String unitId) {
        return store.find(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
    }

    private void verifyCheckpoint(CheckpointRow checkpoint) {
        if (!checksum(checkpoint.snapshotJson()).equals(checkpoint.checksum())) {
            throw new ServiceException("STORAGE_CORRUPT", 409, "Checksum mismatch for " + checkpoint.unitId() + " revision " + checkpoint.revision() + "; restore a verified backup");
        }
    }

    private boolean shouldCheckpoint(WorkbookRow row, OperationEnvelope operation, String envelopeJson) {
        if (row.snapshotRevision() == row.revision()) return false;
        List<OperationRow> sinceCheckpoint = store.listOperationsBetween(row.unitId(), row.snapshotRevision(), row.revision());
        long operationCount = sinceCheckpoint.size();
        long bytes = sinceCheckpoint.stream().mapToLong(entry -> entry.envelopeJson().getBytes(StandardCharsets.UTF_8).length).sum();
        return operationCount + 1 >= CHECKPOINT_OPERATION_LIMIT || bytes + envelopeJson.getBytes(StandardCharsets.UTF_8).length >= CHECKPOINT_BYTES_LIMIT;
    }

    private void audit(String operationId, String unitId, String actor, String eventType, String outcome, String reason, JsonNode details) {
        store.insertAudit(new AuditRecord(UUID.randomUUID(), operationId, unitId, actor, eventType, outcome, reason, details, Instant.now()));
    }

    private void enqueueRevisionEvent(String unitId, String operationId, long revision, String envelopeJson, Instant createdAt) {
        if (!coordination.redisEnabled()) return;
        store.insertOutbox(new OutboxRow(UUID.randomUUID(), unitId, operationId, revision, envelopeJson, createdAt, 0));
    }

    private WorkbookSnapshotResponse response(String unitId, JsonNode snapshot, long revision, String checksum) {
        return new WorkbookSnapshotResponse(unitId, snapshot.deepCopy(), revision, checksum);
    }

    private JsonNode readJson(String json) {
        try {
            return mapper.readTree(json);
        } catch (Exception error) {
            throw new IllegalStateException("Stored snapshot JSON is invalid", error);
        }
    }

    private String writeJson(Object json) {
        try {
            return mapper.writeValueAsString(json);
        } catch (Exception error) {
            throw new IllegalStateException("Unable to serialize workbook JSON", error);
        }
    }

    private String checksum(String json) {
        try {
            return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(json.getBytes(StandardCharsets.UTF_8)));
        } catch (Exception error) {
            throw new IllegalStateException("Unable to checksum workbook snapshot", error);
        }
    }

    public record CommitResult(CommittedOperationEnvelope operation, boolean committed) {
    }

    public record RestoreResult(CommittedOperationEnvelope operation, WorkbookSnapshotResponse snapshot) {
    }
}
