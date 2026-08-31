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
import com.xc.luckysheet.server.store.OperationTail;
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

    public WorkbookSnapshotResponse readSnapshot(String unitId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookRow row = requireWorkbook(unitId);
        JsonNode snapshot = readCurrentSnapshot(row).snapshot();
        if (snapshot.isObject()) ((ObjectNode) snapshot).put("name", row.name());
        String json = writeJson(snapshot);
        return response(unitId, snapshot, row.revision(), checksum(json));
    }

    @Transactional
    public CommitResult commit(String routeUnitId, OperationEnvelope operation, String actor) {
        try {
            if (operation == null) throw ServiceException.validation("Operation is required");
            return commitInternal(routeUnitId, operation, actor);
        } catch (ServiceException error) {
            auditRecorder.rejected(operation == null ? null : operation.operationId(), routeUnitId, actor, "OPERATION_COMMIT", error.getMessage());
            throw error;
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
            return new CommitResult(readCommitted(existing), false);
        }
        OperationRow sequenceExisting = store.findOperationBySequence(routeUnitId, actor, operation.clientSequence()).orElse(null);
        if (sequenceExisting != null && !sequenceExisting.operationId().equals(operation.operationId())) {
            throw ServiceException.conflict("clientSequence was already committed");
        }
        validateIntent(routeUnitId, operation, actor, row);
        if (registry.requiresExactBase(operation.mutations()) && row.revision() != operation.baseRevision()) {
            throw ServiceException.conflict("Revision conflict; rebase against revision " + row.revision());
        }

        SnapshotState current = readCurrentSnapshot(row);
        JsonNode before = current.snapshot();
        JsonNode next = before;
        List<CommittedOperationMutation> committedMutations = new ArrayList<>();
        for (OperationMutation mutation : operation.mutations()) {
            MutationPreparation prepared = registry.prepare(next, mutation, actorRole);
            next = prepared.descriptor().apply(next, mutation);
            validatePhysicalDataBlocks(routeUnitId, mutation);
            committedMutations.add(CommittedOperationMutation.from(mutation, prepared.affectedRanges()));
        }

        long nextRevision = incrementRevision(row.revision());
        Instant committedAt = Instant.now();
        CommittedOperationEnvelope committed = CommittedOperationEnvelope.from(operation, actor, nextRevision, committedAt, committedMutations);
        String envelopeJson = writeJson(committed);
        store.insertOperation(new OperationRow(operation.operationId(), routeUnitId, nextRevision, actor, operation.clientSequence(), operation.baseRevision(), envelopeJson, committedAt));
        enqueueRevisionEvent(routeUnitId, operation.operationId(), nextRevision, envelopeJson, committedAt);
        String canonicalName = next.path("name").asText(row.name()).trim();
        store.updateWorkbookRevisionAndName(routeUnitId, nextRevision, canonicalName, committedAt);
        if (shouldCheckpoint(row, current.tail(), envelopeJson)) {
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
        WorkbookRow current = requireWorkbook(unitId);
        long toExclusive = beforeRevision == Long.MAX_VALUE
                ? incrementRevision(current.revision())
                : Math.min(beforeRevision, incrementRevision(current.revision()));
        long fromInclusive = Math.max(1, toExclusive - Math.max(1, limit));
        List<OperationRow> page = store.readRevisionPage(unitId, fromInclusive, toExclusive, limit);
        validateRevisionPage(page, fromInclusive, toExclusive);
        List<RevisionRecord> items = page.stream().map(this::revisionRecord).toList();
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
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        WorkbookRow row = store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
        if (row.lifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash and cannot be checkpointed");
        SnapshotState current = readCurrentSnapshot(row);
        JsonNode snapshot = current.snapshot();
        String json = writeJson(snapshot);
        if (row.snapshotRevision() == row.revision()) {
            return new CheckpointResponse(response(unitId, snapshot, row.revision(), checksum(json)), false);
        }
        Instant now = Instant.now();
        store.updateWorkbook(unitId, row.revision(), json, row.revision(), now);
        store.insertCheckpoint(unitId, row.revision(), json, checksum(json), now);
        return new CheckpointResponse(response(unitId, snapshot, row.revision(), checksum(json)), true);
    }

    @Transactional
    public RestoreResult restore(String unitId, RestoreRequest request, String actor) {
        access.require(unitId, actor, WorkbookAclRole.OWNER);
        WorkbookRow row = store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
        if (request.targetRevision() > row.revision()) throw ServiceException.notFound("Revision not found: " + request.targetRevision());
        JsonNode target = snapshotAtRevision(row, request.targetRevision());
        WorkbookSnapshotValidator.requireCanonical(target, unitId);
        registry.require("workbook.restore", true);
        long revision = incrementRevision(row.revision());
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
        OperationEnvelope source = new OperationEnvelope(OperationEnvelope.SCHEMA, operationId, unitId, revision, row.revision(),
                sourceMutations, now);
        CommittedOperationEnvelope committed = CommittedOperationEnvelope.system(source, actor, revision, now, mutations);
        String json = writeJson(target);
        String envelopeJson = writeJson(committed);
        // System restores must not share an authenticated browser's client-sequence namespace.
        store.insertOperation(new OperationRow(operationId, unitId, revision, SYSTEM_RESTORE_ACTOR, revision, row.revision(), envelopeJson, now));
        enqueueRevisionEvent(unitId, operationId, revision, envelopeJson, now);
        store.updateWorkbook(unitId, revision, json, revision, now);
        store.insertCheckpoint(unitId, revision, json, checksum(json), now);
        audit(operationId, unitId, actor, "SNAPSHOT_RESTORE", "ACCEPTED", request.reason(), mapper.createObjectNode().put("targetRevision", request.targetRevision()));
        return new RestoreResult(committed, response(unitId, target, revision, checksum(json)));
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
            CommittedOperationEnvelope committed = mapper.readValue(row.envelopeJson(), CommittedOperationEnvelope.class);
            if (!row.operationId().equals(committed.operationId())
                    || !row.unitId().equals(committed.unitId())
                    || row.revision() != committed.revision()) {
                throw historyGap("Stored operation identity does not match its revision row", null);
            }
            return committed;
        } catch (Exception error) {
            if (error instanceof ServiceException serviceException) throw serviceException;
            throw historyGap("Stored operation envelope is invalid", error);
        }
    }

    private SnapshotState readCurrentSnapshot(WorkbookRow row) {
        if (row.snapshotRevision() < 0 || row.revision() < row.snapshotRevision()) {
            throw historyGap("Workbook snapshot revision is outside the current revision", null);
        }
        JsonNode snapshot = WorkbookSnapshotValidator.migrateStored(readJson(row.snapshotJson()), row.unitId());
        if (row.snapshotRevision() == row.revision()) return new SnapshotState(snapshot, OperationTail.empty(row.unitId(), row.revision(), row.revision()));
        OperationTail tail = store.readOperationTail(row.unitId(), row.snapshotRevision(), row.revision());
        return new SnapshotState(replay(snapshot, tail), tail);
    }

    private JsonNode snapshotAtRevision(WorkbookRow current, long targetRevision) {
        if (targetRevision < 0 || targetRevision > current.revision()) throw ServiceException.notFound("Revision not found: " + targetRevision);
        if (targetRevision == current.revision()) return readCurrentSnapshot(current).snapshot();
        CheckpointRow checkpoint = store.findCheckpoint(current.unitId(), targetRevision)
                .orElseGet(() -> store.findLatestCheckpointAtOrBefore(current.unitId(), targetRevision)
                        .orElseThrow(() -> historyGap("No checkpoint covers revision " + targetRevision, null)));
        if (checkpoint.revision() < 0 || checkpoint.revision() > targetRevision) {
            throw historyGap("Checkpoint revision is outside the requested interval", null);
        }
        JsonNode snapshot = checkpointSnapshot(checkpoint, current.unitId());
        if (checkpoint.revision() == targetRevision) return snapshot;
        OperationTail tail = store.readOperationTail(current.unitId(), checkpoint.revision(), targetRevision);
        return replay(snapshot, tail);
    }

    private WorkbookRow requireWorkbook(String unitId) {
        return store.find(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
    }

    private boolean shouldCheckpoint(WorkbookRow row, OperationTail tail, String envelopeJson) {
        if (row.snapshotRevision() == row.revision()) return false;
        long operationCount = tail.operations().size() + 1L;
        long bytes = tail.envelopeBytes() + envelopeJson.getBytes(StandardCharsets.UTF_8).length;
        return operationCount >= CHECKPOINT_OPERATION_LIMIT || bytes >= CHECKPOINT_BYTES_LIMIT;
    }

    /** Replays exactly the interval carried by the tail and nothing else. */
    private JsonNode replay(JsonNode snapshot, OperationTail tail) {
        long expected = tail.fromExclusive() + 1;
        for (OperationRow operation : tail.operations()) {
            if (operation.revision() != expected || operation.revision() > tail.toInclusive()) {
                throw historyGap("Operation history is not contiguous at revision " + expected, null);
            }
            CommittedOperationEnvelope committed = readCommitted(operation);
            if (committed.mutations().stream().anyMatch(mutation -> "workbook.restore".equals(mutation.id()))) {
                throw historyGap("Restore revision " + operation.revision() + " has no covering checkpoint", null);
            }
            List<OperationMutation> mutations = committed.mutations().stream()
                    .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList();
            snapshot = registry.applyPublicMutations(snapshot, mutations);
            expected++;
        }
        if (expected != tail.toInclusive() + 1) {
            throw historyGap("Operation history is missing a revision in " + (tail.fromExclusive() + 1) + ".." + tail.toInclusive(), null);
        }
        return snapshot;
    }

    private JsonNode checkpointSnapshot(CheckpointRow checkpoint, String unitId) {
        if (checkpoint.snapshotJson() == null || checkpoint.checksum() == null) {
            throw historyGap("Checkpoint payload is incomplete at revision " + checkpoint.revision(), null);
        }
        String actual = checksum(checkpoint.snapshotJson());
        if (!actual.equalsIgnoreCase(checkpoint.checksum())) {
            throw historyGap("Checkpoint checksum does not match revision " + checkpoint.revision(), null);
        }
        return WorkbookSnapshotValidator.migrateStored(readJson(checkpoint.snapshotJson()), unitId);
    }

    private void validateRevisionPage(List<OperationRow> page, long fromInclusive, long toExclusive) {
        long expected = fromInclusive;
        for (OperationRow operation : page.stream().sorted(java.util.Comparator.comparingLong(OperationRow::revision)).toList()) {
            if (operation.revision() != expected || operation.revision() >= toExclusive) {
                throw historyGap("Revision log is not contiguous at revision " + expected, null);
            }
            readCommitted(operation);
            expected++;
        }
        if (!page.isEmpty() && expected != fromInclusive + page.size()) {
            throw historyGap("Revision log contains a missing revision", null);
        }
    }

    private void validatePhysicalDataBlocks(String unitId, OperationMutation mutation) {
        String id = mutation.id();
        if (!(id.equals("dataSource.add") || id.equals("dataSource.update") || id.startsWith("query.load."))) return;
        JsonNode params = mutation.params();
        if (params == null || !params.isObject()) throw ServiceException.validation("Mutation params must be an object");
        JsonNode source = params.get("source");
        if (source == null || source.isNull()) return;
        if (!source.isObject()) throw ServiceException.validation("Data source manifest is required");
        store.validateDataSourceManifestBlocks(unitId, (ObjectNode) source);
    }

    private long incrementRevision(long revision) {
        if (revision == Long.MAX_VALUE) throw historyGap("Workbook revision overflowed the canonical range", null);
        return revision + 1;
    }

    private ServiceException historyGap(String message, Throwable cause) {
        String detail = message.startsWith("HISTORY_GAP") ? message : "HISTORY_GAP: " + message;
        return cause == null ? new ServiceException("CONFLICT", 409, detail) : new ServiceException("CONFLICT", 409, detail, cause);
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
            JsonNode parsed = mapper.readTree(json);
            if (parsed == null) throw new IllegalStateException("empty JSON");
            return parsed;
        } catch (Exception error) {
            if (error instanceof ServiceException serviceException) throw serviceException;
            throw historyGap("Stored snapshot JSON is invalid", error);
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

    private record SnapshotState(JsonNode snapshot, OperationTail tail) {
    }
}
