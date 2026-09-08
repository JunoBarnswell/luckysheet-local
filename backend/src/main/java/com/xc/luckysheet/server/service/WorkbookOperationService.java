package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.*;
import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.store.*;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.support.TransactionSynchronization;
import org.springframework.transaction.support.TransactionSynchronizationManager;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Instant;
import java.util.HexFormat;
import java.util.List;
import java.util.UUID;

/** Authentication and atomic persistence boundary. Rust owns all workbook semantics. */
@Service
public class WorkbookOperationService {
    private final WorkbookStore store;
    private final AccessControlService access;
    private final ObjectMapper mapper;
    private final AuditRecorder auditRecorder;
    private final CoordinationProperties coordination;
    private final KernelHostClient kernel;
    private final KernelPersistenceService persistence;
    private final QueryExecutionProofService queryProofs;

    public WorkbookOperationService(WorkbookStore store, AccessControlService access, ObjectMapper mapper,
            AuditRecorder auditRecorder, CoordinationProperties coordination, KernelHostClient kernel,
            KernelPersistenceService persistence, QueryExecutionProofService queryProofs) {
        this.store = store; this.access = access; this.mapper = mapper; this.auditRecorder = auditRecorder;
        this.coordination = coordination; this.kernel = kernel; this.persistence = persistence; this.queryProofs = queryProofs;
    }

    public WorkbookOpenResponse open(String unitId, String actor) {
        WorkbookRow row = readable(unitId, actor);
        return response(unitId, row.revision(), persistence.readManifest(unitId, row.revision()));
    }

    public JsonNode manifest(String unitId, Long revision, String actor) {
        WorkbookRow row = readable(unitId, actor);
        long requested = revision == null ? row.revision() : revision;
        if (requested < 0 || requested > row.revision()) throw ServiceException.notFound("Revision not found: " + requested);
        return persistence.readManifest(unitId, requested);
    }

    public JsonNode page(String unitId, long revision, String sheetId, int pageRow, int pageColumn, String actor) {
        WorkbookRow row = readable(unitId, actor);
        if (revision < 0 || revision > row.revision()) throw ServiceException.notFound("Revision not found: " + revision);
        return persistence.readPage(unitId, revision, sheetId, pageRow, pageColumn);
    }

    @Transactional
    public CommitResult commit(String unitId, OperationEnvelope operation, String actor) {
        try {
            if (operation == null || !unitId.equals(operation.unitId())) throw ServiceException.validation("Operation unitId must match route");
            WorkbookAclRole role = access.require(unitId, actor, WorkbookAclRole.VIEWER);
            WorkbookRow row = lockedActive(unitId);
            OperationRow existing = store.findOperation(operation.operationId()).orElse(null);
            if (existing != null) {
                if (!actor.equals(existing.actorSubject()) || !unitId.equals(existing.unitId()))
                    throw ServiceException.forbidden("Operation belongs to another subject or workbook");
                return new CommitResult(readCommitted(existing), false, persistence.readChangeSet(unitId, operation.operationId()));
            }
            OperationRow sequence = store.findOperationBySequence(unitId, actor, operation.clientSequence()).orElse(null);
            if (sequence != null) throw ServiceException.conflict("clientSequence was already committed");
            if (operation.baseRevision() != row.revision()) throw ServiceException.conflict("Revision conflict; current revision is " + row.revision());
            validateIntent(unitId, operation, actor, row);
            ObjectNode command = mapper.createObjectNode().put("unitId", unitId).put("baseRevision", row.revision())
                    .put("operationId", operation.operationId()).put("accessRole", role.wireValue());
            ObjectNode commandParams = mapper.createObjectNode();
            if (operation.intent() != null) {
                command.put("commandId", "history.undo");
                commandParams.set("history", persistence.historyForUndo(unitId, operation.intent().targetOperationId()));
            } else {
                command.put("commandId", "operation.apply");
                commandParams.set("mutations", mapper.valueToTree(operation.mutations()));
                validateExternalData(unitId, row.revision(), actor, commandParams);
            }
            command.set("params", commandParams);
            JsonNode result;
            synchronized (kernel) {
                invalidateOnRollback();
                persistence.reopen(unitId, row.revision(), kernel);
                hydrateCommand(command, unitId, row.revision());
                result = kernel.call("command", command);
            }
            return persistCommit(row, operation, actor, result);
        } catch (ServiceException error) {
            auditRecorder.rejected(operation == null ? null : operation.operationId(), unitId, actor, "OPERATION_COMMIT", error.getMessage());
            throw error;
        }
    }

    /**
     * Converts a server-owned semantic intent into the same canonical operation
     * envelope used by editor clients. The workbook lock serializes sequence
     * allocation with the commit, so catalog actions cannot create a parallel
     * metadata write path.
     */
    @Transactional
    public CommitResult commitServerMutation(String unitId, OperationMutation mutation, String actor, String operationPrefix) {
        WorkbookRow row = lockedActive(unitId);
        long sequence = store.nextClientSequence(unitId, actor);
        OperationEnvelope operation = new OperationEnvelope(
                OperationEnvelope.SCHEMA,
                operationPrefix + "-" + UUID.randomUUID(),
                unitId,
                sequence,
                row.revision(),
                List.of(mutation),
                Instant.now()
        );
        return commit(unitId, operation, actor);
    }

    private void hydrateCommand(ObjectNode command, String unitId, long revision) {
        JsonNode plan = kernel.call("command.prepare", command);
        if (!plan.path("pages").isArray()) throw new KernelHostException("PROTOCOL_ERROR", "Command preparation has no page plan", unitId, "deploy-matching-kernel-host");
        for (JsonNode key : plan.path("pages")) {
            persistence.loadPage(unitId, revision, key.path("sheetId").asText(), key.path("pageRow").asInt(-1), key.path("pageColumn").asInt(-1), kernel);
        }
    }

    private CommitResult persistCommit(WorkbookRow row, OperationEnvelope source, String actor, JsonNode result) {
        long revision = result.path("revision").asLong(-1);
        if (revision != row.revision() + 1 || result.path("baseRevision").asLong(-1) != row.revision()
                || !source.operationId().equals(result.path("operationId").asText()))
            throw new KernelHostException("REVISION_INVALID", "Native changeset identity/revision is not contiguous", row.unitId(), "reopen-last-committed-manifest");
        List<RangeRef> affected;
        try { affected = mapper.convertValue(result.path("affectedRanges"), new com.fasterxml.jackson.core.type.TypeReference<List<RangeRef>>() {}); }
        catch (IllegalArgumentException error) { throw new KernelHostException("PROTOCOL_ERROR", "Native affected ranges are invalid", row.unitId(), "deploy-matching-kernel-host"); }
        List<CommittedOperationMutation> mutations = source.intent() != null
                ? List.of(new CommittedOperationMutation("history.undo", "workbook", mapper.createObjectNode().put("targetOperationId", source.intent().targetOperationId()), affected))
                : source.mutations().stream().map(m -> new CommittedOperationMutation(m.id(), m.sheetId(), m.params(), affected)).toList();
        Instant now = Instant.now();
        CommittedOperationEnvelope committed = CommittedOperationEnvelope.from(source, actor, revision, now, mutations);
        String json = writeJson(committed);
        persistence.publish(result, row.unitId(), revision);
        store.insertOperation(new OperationRow(source.operationId(), row.unitId(), revision, actor, source.clientSequence(), row.revision(), json, now));
        // The durable log is always written. Redis is only a delivery transport.
        store.insertOutbox(new OutboxRow(UUID.randomUUID(), row.unitId(), source.operationId(), revision, json, now, 0));
        store.updateWorkbookRevisionAndName(row.unitId(), revision, result.path("manifest").path("name").asText(), now);
        audit(source.operationId(), row.unitId(), actor, "OPERATION_COMMIT", null, mapper.createObjectNode().put("revision", revision));
        // Use the durable reconstruction for both the first acknowledgement and
        // idempotent replay. History undo can restore a content-addressed page
        // without asking native to re-emit its bytes; readChangeSet resolves the
        // verified payload so the browser replica never receives a manifest-only
        // changed page.
        return new CommitResult(committed, true, persistence.readChangeSet(row.unitId(), source.operationId()));
    }

    private void validateExternalData(String unitId, long revision, String actor, ObjectNode commandParams) {
        for (JsonNode mutation : commandParams.path("mutations")) {
            String id = mutation.path("id").asText();
            JsonNode params = mutation.path("params");
            if (id.startsWith("query.load.")) {
                if (!params.path("sourceRevision").isIntegralNumber() || params.path("sourceRevision").asLong(-1) != revision)
                    throw ServiceException.conflict("Query load must carry the exact sourceRevision");
                JsonNode proof = queryProofs.consumeProof(unitId, params.path("queryId").asText(), params.path("executionToken").asText(),
                        actor, revision, params.path("resultHash").asText());
                ((ObjectNode) params).set("trustedResult", proof);
            }
            if (id.equals("dataSource.add") || id.equals("dataSource.update") || id.startsWith("query.load.")) {
                JsonNode dataSource = params.path("source");
                if (dataSource.isObject()) store.validateDataSourceManifestBlocks(unitId, (ObjectNode) dataSource);
            }
        }
    }

    private void validateIntent(String unitId, OperationEnvelope operation, String actor, WorkbookRow row) {
        OperationIntent intent = operation.intent();
        if (intent == null) return;
        if (!OperationIntent.UNDO.equals(intent.type())) throw ServiceException.validation("Unsupported operation intent");
        if (intent.targetOperationId().equals(operation.operationId())) throw ServiceException.conflict("Undo cannot target itself");
        OperationRow target = store.findOperation(intent.targetOperationId()).orElseThrow(() -> ServiceException.conflict("Undo target is not committed"));
        if (!unitId.equals(target.unitId()) || !actor.equals(target.actorSubject())) throw ServiceException.forbidden("Undo target belongs to another workbook or subject");
        if (intent.targetBaseRevision() != target.baseRevision()) throw ServiceException.conflict("Undo target revision does not match");
    }

    public CursorPage<RevisionRecord> revisions(String unitId, String actor, long beforeRevision, int limit, String cursor) {
        WorkbookRow row = readable(unitId, actor);
        int bounded = Math.max(1, Math.min(200, limit));
        long end = Math.min(beforeRevision, Math.addExact(row.revision(), 1));
        long start = Math.max(1, end - bounded);
        List<OperationRow> page = store.readRevisionPage(unitId, start, end, bounded);
        if (page.size() != end - start) throw new ServiceException("HISTORY_GAP", 409, "Revision log is incomplete");
        List<RevisionRecord> items = page.stream().map(op -> new RevisionRecord(op.operationId(), op.revision(), op.committedAt(), readCommitted(op))).toList();
        return new CursorPage<>(items, items.size() == bounded ? Long.toString(start) : null);
    }

    public WorkbookOpenResponse readRevision(String unitId, long revision, String actor) {
        return response(unitId, revision, manifest(unitId, revision, actor));
    }

    @Transactional
    public CheckpointResponse checkpoint(String unitId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        WorkbookRow row = lockedActive(unitId);
        // Every successful page transaction already publishes its immutable manifest checkpoint.
        return new CheckpointResponse(response(unitId, row.revision(), persistence.readManifest(unitId, row.revision())), false);
    }

    @Transactional
    public RestoreResult restore(String unitId, RestoreRequest request, String actor) {
        access.require(unitId, actor, WorkbookAclRole.OWNER);
        WorkbookRow row = lockedActive(unitId);
        if (request.targetRevision() < 0 || request.targetRevision() > row.revision()) throw ServiceException.notFound("Restore revision not found");
        JsonNode target = persistence.readManifest(unitId, request.targetRevision());
        String operationId = UUID.randomUUID().toString();
        ObjectNode params = mapper.createObjectNode().put("unitId", unitId).put("baseRevision", row.revision()).put("operationId", operationId).put("accessRole", "owner");
        params.set("targetManifest", target);
        JsonNode result;
        synchronized (kernel) {
            invalidateOnRollback();
            persistence.reopen(unitId, row.revision(), kernel);
            result = kernel.call("restore", params);
        }
        Instant now = Instant.now();
        ObjectNode sourceParams = mapper.createObjectNode().put("targetRevision", request.targetRevision()).put("reason", request.reason());
        OperationEnvelope source = new OperationEnvelope(OperationEnvelope.SCHEMA, operationId, unitId, row.revision() + 1,
                row.revision(), List.of(new OperationMutation("workbook.restore", "workbook", sourceParams)), now);
        CommitResult committed = persistCommit(row, source, "system:restore:" + actor, result);
        return new RestoreResult(committed.operation(), response(unitId, committed.operation().revision(), result.path("manifest")), result);
    }

    public List<AclEntry> acl(String unitId, String actor) { return access.list(unitId, actor); }
    public WorkbookAccessProjection accessProjection(String unitId, String actor) {
        return new WorkbookAccessProjection(unitId, access.currentRole(unitId, actor), store.nextClientSequence(unitId, actor));
    }
    public AclEntry grantAcl(String unitId, String actor, String target, WorkbookAclRole role) { return access.grant(unitId, actor, target, role); }
    public void revokeAcl(String unitId, String actor, String target) { access.revoke(unitId, actor, target); }
    public List<AuditRecord> audit(String unitId, String actor, int limit) { readable(unitId, actor); return store.listAudit(unitId, Math.min(200, Math.max(1, limit))); }

    private WorkbookRow readable(String unitId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        return store.find(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
    }
    private WorkbookRow lockedActive(String unitId) {
        WorkbookRow row = store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
        if (row.lifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash");
        return row;
    }
    private void invalidateOnRollback() {
        if (!TransactionSynchronizationManager.isSynchronizationActive()) throw new IllegalStateException("Native mutation requires an active database transaction");
        TransactionSynchronizationManager.registerSynchronization(new TransactionSynchronization() {
            @Override public void afterCompletion(int status) { if (status != STATUS_COMMITTED) kernel.abortTransaction(); }
        });
    }
    private CommittedOperationEnvelope readCommitted(OperationRow row) {
        try {
            CommittedOperationEnvelope result = mapper.readValue(row.envelopeJson(), CommittedOperationEnvelope.class);
            if (!row.operationId().equals(result.operationId()) || !row.unitId().equals(result.unitId()) || row.revision() != result.revision())
                throw new IllegalArgumentException("Stored identity mismatch");
            return result;
        } catch (Exception error) { throw new ServiceException("HISTORY_GAP", 409, "Stored committed operation is invalid", error); }
    }
    private void audit(String operationId, String unitId, String actor, String event, String reason, JsonNode details) {
        store.insertAudit(new AuditRecord(UUID.randomUUID(), operationId, unitId, actor, event, "ACCEPTED", reason, details, Instant.now()));
    }
    private WorkbookOpenResponse response(String unitId, long revision, JsonNode manifest) {
        try { return new WorkbookOpenResponse(unitId, revision, manifest, HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(writeJson(manifest).getBytes(StandardCharsets.UTF_8)))); }
        catch (java.security.NoSuchAlgorithmException error) { throw new IllegalStateException(error); }
    }
    private String writeJson(Object value) { try { return mapper.writeValueAsString(value); } catch (Exception error) { throw new IllegalStateException("Unable to encode canonical result", error); } }
    public record CommitResult(CommittedOperationEnvelope operation, boolean committed, JsonNode changeSet) { }
    public record RestoreResult(CommittedOperationEnvelope operation, WorkbookOpenResponse workbook, JsonNode changeSet) { }
}
