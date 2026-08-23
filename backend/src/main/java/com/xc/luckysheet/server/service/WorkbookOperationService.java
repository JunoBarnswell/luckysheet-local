package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.AclEntry;
import com.xc.luckysheet.server.contract.AuditRecord;
import com.xc.luckysheet.server.contract.CheckpointResponse;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.RestoreRequest;
import com.xc.luckysheet.server.contract.RevisionRecord;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookSnapshotResponse;
import com.xc.luckysheet.server.contract.WorkbookSummary;
import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.mutation.MutationDescriptorRegistry;
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

@Service
public class WorkbookOperationService {
    private static final long CHECKPOINT_OPERATION_LIMIT = 50;
    private static final long CHECKPOINT_BYTES_LIMIT = 512_000;

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

    @Transactional
    public WorkbookSnapshotResponse create(CreateWorkbookRequest request, String actor) {
        if (store.find(request.unitId()).isPresent()) throw ServiceException.conflict("Workbook already exists");
        Instant now = Instant.now();
        String snapshotJson = writeJson(request.snapshot());
        store.insertWorkbook(request.unitId(), request.name().trim(), snapshotJson, now);
        store.insertAcl(request.unitId(), actor, WorkbookAclRole.OWNER, now);
        store.insertCheckpoint(request.unitId(), 0, snapshotJson, checksum(snapshotJson), now);
        return response(request.unitId(), request.snapshot(), 0, checksum(snapshotJson));
    }

    public List<WorkbookSummary> list(String actor) {
        return store.listForSubject(actor);
    }

    public WorkbookSnapshotResponse readSnapshot(String unitId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookRow row = requireWorkbook(unitId);
        JsonNode snapshot = currentSnapshot(row);
        String json = writeJson(snapshot);
        return response(unitId, snapshot, row.revision(), checksum(json));
    }

    @Transactional
    public CommitResult commit(String routeUnitId, OperationEnvelope operation, String actor) {
        try {
            return commitInternal(routeUnitId, operation, actor);
        } catch (ServiceException error) {
            auditRecorder.rejected(operation == null ? null : operation.operationId(), routeUnitId, actor, "OPERATION_COMMIT", error.getMessage());
            throw error;
        }
    }

    private CommitResult commitInternal(String routeUnitId, OperationEnvelope operation, String actor) {
        if (!routeUnitId.equals(operation.unitId())) throw ServiceException.validation("Operation unitId does not match route");
        access.require(routeUnitId, actor, WorkbookAclRole.EDITOR);
        WorkbookRow row = store.findForUpdate(routeUnitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + routeUnitId));

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
        if (row.revision() != operation.baseRevision()) {
            throw ServiceException.conflict("Revision conflict; rebase against revision " + row.revision());
        }

        JsonNode before = currentSnapshot(row);
        JsonNode next = before;
        List<CommittedOperationMutation> committedMutations = new ArrayList<>();
        for (OperationMutation mutation : operation.mutations()) {
            List<RangeRef> ranges = registry.resolveRanges(next, mutation);
            next = registry.applyPublicMutations(next, List.of(mutation));
            committedMutations.add(CommittedOperationMutation.from(mutation, ranges));
        }

        long nextRevision = row.revision() + 1;
        Instant committedAt = Instant.now();
        CommittedOperationEnvelope committed = CommittedOperationEnvelope.from(operation, actor, nextRevision, committedAt, committedMutations);
        String envelopeJson = writeJson(committed);
        store.insertOperation(new OperationRow(operation.operationId(), routeUnitId, nextRevision, actor, operation.clientSequence(), operation.baseRevision(), envelopeJson, committedAt));
        enqueueRevisionEvent(routeUnitId, operation.operationId(), nextRevision, envelopeJson, committedAt);
        store.updateWorkbookRevision(routeUnitId, nextRevision, committedAt);
        if (shouldCheckpoint(row, operation, envelopeJson)) {
            String nextJson = writeJson(next);
            store.updateWorkbook(routeUnitId, nextRevision, nextJson, nextRevision, committedAt);
            store.insertCheckpoint(routeUnitId, nextRevision, nextJson, checksum(nextJson), committedAt);
        }
        audit(operation.operationId(), routeUnitId, actor, "OPERATION_COMMIT", "ACCEPTED", null, mapper.createObjectNode().put("revision", nextRevision));
        return new CommitResult(committed, true);
    }

    public List<RevisionRecord> revisions(String unitId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        return store.listOperations(unitId).stream().map(this::revisionRecord).toList();
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
        if (row.snapshotRevision() == row.revision()) {
            JsonNode snapshot = readJson(row.snapshotJson());
            return new CheckpointResponse(response(unitId, snapshot, row.revision(), checksum(row.snapshotJson())), false);
        }
        JsonNode snapshot = currentSnapshot(row);
        String json = writeJson(snapshot);
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
        OperationEnvelope source = new OperationEnvelope(OperationEnvelope.SCHEMA, operationId, unitId, revision, row.revision(),
                sourceMutations, now);
        CommittedOperationEnvelope committed = CommittedOperationEnvelope.from(source, actor, revision, now, mutations);
        String json = writeJson(target);
        String envelopeJson = writeJson(committed);
        store.insertOperation(new OperationRow(operationId, unitId, revision, actor, revision, row.revision(), envelopeJson, now));
        enqueueRevisionEvent(unitId, operationId, revision, envelopeJson, now);
        store.updateWorkbook(unitId, revision, json, revision, now);
        store.insertCheckpoint(unitId, revision, json, checksum(json), now);
        audit(operationId, unitId, actor, "SNAPSHOT_RESTORE", "ACCEPTED", request.reason(), mapper.createObjectNode().put("targetRevision", request.targetRevision()));
        return new RestoreResult(committed, response(unitId, target, revision, checksum(json)));
    }

    public List<AclEntry> acl(String unitId, String actor) {
        return access.list(unitId, actor);
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
        JsonNode snapshot = readJson(row.snapshotJson());
        if (row.snapshotRevision() == row.revision()) return snapshot;
        for (OperationRow operation : store.listOperations(row.unitId()).stream()
                .filter(entry -> entry.revision() > row.snapshotRevision() && entry.revision() <= row.revision())
                .sorted(java.util.Comparator.comparingLong(OperationRow::revision)).toList()) {
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
        JsonNode snapshot = readJson(checkpoint.snapshotJson());
        if (checkpoint.revision() == targetRevision) return snapshot;
        for (OperationRow operation : store.listOperations(current.unitId()).stream()
                .filter(entry -> entry.revision() > checkpoint.revision() && entry.revision() <= targetRevision)
                .sorted(java.util.Comparator.comparingLong(OperationRow::revision)).toList()) {
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

    private boolean shouldCheckpoint(WorkbookRow row, OperationEnvelope operation, String envelopeJson) {
        if (row.snapshotRevision() == row.revision()) return false;
        long operationCount = store.listOperations(row.unitId()).stream().filter(entry -> entry.revision() > row.snapshotRevision()).count();
        long bytes = store.listOperations(row.unitId()).stream().filter(entry -> entry.revision() > row.snapshotRevision()).mapToLong(entry -> entry.envelopeJson().getBytes(StandardCharsets.UTF_8).length).sum();
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
