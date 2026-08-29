package com.xc.luckysheet.server.store;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.AclEntry;
import com.xc.luckysheet.server.contract.AuditRecord;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.persistence.AuditEntity;
import com.xc.luckysheet.server.persistence.AuditEntityRepository;
import com.xc.luckysheet.server.persistence.CheckpointEntity;
import com.xc.luckysheet.server.persistence.CheckpointEntityRepository;
import com.xc.luckysheet.server.persistence.DataBlockEntity;
import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import com.xc.luckysheet.server.persistence.OperationEntity;
import com.xc.luckysheet.server.persistence.OperationEntityRepository;
import com.xc.luckysheet.server.persistence.OutboxEntity;
import com.xc.luckysheet.server.persistence.OutboxEntityRepository;
import com.xc.luckysheet.server.persistence.ShareEntity;
import com.xc.luckysheet.server.persistence.ShareEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookAclEntity;
import com.xc.luckysheet.server.persistence.WorkbookAclEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.service.ServiceException;
import org.springframework.data.domain.PageRequest;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.transaction.annotation.Propagation;

import java.time.Instant;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.UUID;
import java.util.HexFormat;
import com.fasterxml.jackson.databind.node.ObjectNode;

/**
 * ORM-backed persistence boundary for workbook state.
 *
 * <p>The service-facing record contract remains stable while all database
 * access is performed by Spring Data JPA repositories. JSON is text,
 * timestamps are {@link Instant}, and binary payloads use JPA LOBs so
 * Hibernate can select the correct type for the configured database.</p>
 */
@Repository
public class WorkbookStore {
    private final WorkbookEntityRepository workbooks;
    private final WorkbookAclEntityRepository acl;
    private final ShareEntityRepository shares;
    private final OperationEntityRepository operations;
    private final CheckpointEntityRepository checkpoints;
    private final OutboxEntityRepository outbox;
    private final AuditEntityRepository audits;
    private final DataBlockEntityRepository dataBlocks;
    private final ObjectMapper mapper;

    /** Maximum number of rows materialized by any one replay query. */
    public static final int MAX_OPERATION_TAIL_ENTRIES = 512;
    private static final long MAX_OPERATION_TAIL_BYTES = 8L * 1024L * 1024L;
    private static final int MAX_MANIFEST_BLOCKS = 10_000;

    public WorkbookStore(
            WorkbookEntityRepository workbooks,
            WorkbookAclEntityRepository acl,
            ShareEntityRepository shares,
            OperationEntityRepository operations,
            CheckpointEntityRepository checkpoints,
            OutboxEntityRepository outbox,
            AuditEntityRepository audits,
            DataBlockEntityRepository dataBlocks,
            ObjectMapper mapper
    ) {
        this.workbooks = workbooks;
        this.acl = acl;
        this.shares = shares;
        this.operations = operations;
        this.checkpoints = checkpoints;
        this.outbox = outbox;
        this.audits = audits;
        this.dataBlocks = dataBlocks;
        this.mapper = mapper;
    }

    public Optional<WorkbookRow> find(String unitId) {
        return workbooks.findById(unitId).map(this::workbookRow);
    }

    public Optional<WorkbookRow> findForUpdate(String unitId) {
        return workbooks.findForUpdate(unitId).map(this::workbookRow);
    }

    public void updateWorkbook(String unitId, long revision, String snapshotJson, long snapshotRevision, Instant now) {
        WorkbookEntity entity = workbooks.findById(unitId).orElseThrow(() -> new IllegalStateException("Workbook not found: " + unitId));
        entity.updateSnapshot(revision, snapshotJson, snapshotRevision, now);
        workbooks.save(entity);
    }

    public void updateWorkbookRevision(String unitId, long revision, Instant now) {
        WorkbookEntity entity = workbooks.findById(unitId).orElseThrow(() -> new IllegalStateException("Workbook not found: " + unitId));
        entity.updateRevision(revision, now);
        workbooks.save(entity);
    }

    public void updateWorkbookRevisionAndName(String unitId, long revision, String name, Instant now) {
        WorkbookEntity entity = workbooks.findById(unitId).orElseThrow(() -> new IllegalStateException("Workbook not found: " + unitId));
        entity.updateRevisionAndName(revision, name, now);
        workbooks.save(entity);
    }

    public List<AclEntry> listAcl(String unitId) {
        return acl.findAllForWorkbook(unitId).stream()
                .map(entry -> new AclEntry(entry.getId().getUnitId(), entry.getId().getSubject(), entry.getRole(), entry.getCreatedAt(), entry.getUpdatedAt()))
                .toList();
    }

    public Optional<WorkbookAclRole> findRole(String unitId, String subject) {
        return acl.findForSubject(unitId, subject).map(WorkbookAclEntity::getRole);
    }

    public void insertShare(ShareRow share) {
        shares.save(new ShareEntity(share.shareId(), share.unitId(), share.tokenHash(), share.role(), share.expiresAt(),
                share.revokedAt(), share.createdBy(), share.createdAt()));
    }

    public Optional<ShareRow> findShare(UUID shareId) {
        return shares.findById(shareId).map(this::shareRow);
    }

    public List<ShareRow> listShares(String unitId) {
        return shares.findByUnitIdOrderByCreatedAtDesc(unitId).stream().map(this::shareRow).toList();
    }

    public Optional<ShareRow> findActiveShare(String unitId, UUID shareId, Instant now) {
        return shares.findActiveForWorkbook(unitId, shareId, now).map(this::shareRow);
    }

    public Optional<ShareRow> findActiveShare(UUID shareId, Instant now) {
        return shares.findActive(shareId, now).map(this::shareRow);
    }

    @Transactional
    public int revokeShare(String unitId, UUID shareId, Instant revokedAt) {
        Optional<ShareEntity> entity = shares.findByUnitIdAndShareId(unitId, shareId)
                .filter(share -> share.getRevokedAt() == null);
        if (entity.isEmpty()) return 0;
        entity.get().revoke(revokedAt);
        shares.save(entity.get());
        return 1;
    }

    @Transactional
    public void upsertAcl(String unitId, String subject, WorkbookAclRole role, Instant now) {
        WorkbookAclEntity entity = acl.findForSubject(unitId, subject)
                .orElseGet(() -> new WorkbookAclEntity(unitId, subject, role, now, now));
        entity.updateRole(role, now);
        acl.save(entity);
    }

    public int deleteAcl(String unitId, String subject) {
        Optional<WorkbookAclEntity> entity = acl.findForSubject(unitId, subject);
        if (entity.isEmpty()) return 0;
        acl.delete(entity.get());
        return 1;
    }

    public Optional<OperationRow> findOperation(String operationId) {
        return operations.findById(operationId).map(this::operationRow);
    }

    public Optional<OperationRow> findOperationBySequence(String unitId, String actorSubject, long sequence) {
        return operations.findByUnitIdAndActorSubjectAndClientSequence(unitId, actorSubject, sequence).map(this::operationRow);
    }

    public void insertOperation(OperationRow operation) {
        operations.save(new OperationEntity(operation.operationId(), operation.unitId(), operation.revision(), operation.actorSubject(),
                operation.clientSequence(), operation.baseRevision(), operation.envelopeJson(), operation.committedAt()));
    }

    @Transactional
    public void insertOutbox(OutboxRow event) {
        if (outbox.findByUnitIdAndRevision(event.unitId(), event.revision()).isPresent()) return;
        outbox.save(new OutboxEntity(event.eventId(), event.unitId(), event.operationId(), event.revision(), event.payloadJson(),
                event.createdAt(), event.createdAt(), event.attempts()));
    }

    /** Claims rows with ordinary JPA pessimistic locks, without vendor SQL. */
    @Transactional
    public List<OutboxRow> claimOutbox(Instant now, Instant leaseUntil, int limit) {
        int boundedLimit = Math.max(1, Math.min(limit, 1000));
        List<OutboxEntity> pending = outbox.findPendingForUpdate(now, PageRequest.of(0, boundedLimit));
        pending.forEach(entity -> {
            entity.claim(leaseUntil);
            outbox.save(entity);
        });
        return pending.stream().map(this::outboxRow).toList();
    }

    @Transactional
    public void markOutboxPublished(UUID eventId, Instant publishedAt) {
        outbox.findById(eventId).filter(entity -> entity.getPublishedAt() == null).ifPresent(entity -> {
            entity.markPublished(publishedAt);
            outbox.save(entity);
        });
    }

    @Transactional
    public void releaseOutbox(UUID eventId, Instant nextAttemptAt) {
        outbox.findById(eventId).filter(entity -> entity.getPublishedAt() == null).ifPresent(entity -> {
            entity.release(nextAttemptAt);
            outbox.save(entity);
        });
    }

    @Transactional
    public int deletePublishedOutboxBefore(Instant cutoff) {
        List<OutboxEntity> rows = outbox.findByPublishedAtIsNotNullAndPublishedAtBefore(cutoff);
        if (rows.isEmpty()) return 0;
        outbox.deleteAllInBatch(rows);
        return rows.size();
    }

    /**
     * Reads one bounded revision interval in ascending order.  The extra row
     * is fetched only to prove that the caller did not exceed the replay
     * budget; it is never returned or silently truncated.
     */
    public OperationTail readOperationTail(String unitId, long fromExclusive, long toInclusive) {
        if (unitId == null || unitId.isBlank()) throw ServiceException.validation("Operation tail unitId is required");
        if (fromExclusive < 0 || toInclusive < fromExclusive) throw ServiceException.validation("Operation tail interval is invalid");
        if (fromExclusive == toInclusive) return OperationTail.empty(unitId, fromExclusive, toInclusive);
        List<OperationEntity> entities = operations
                .findByUnitIdAndRevisionGreaterThanAndRevisionLessThanEqualOrderByRevisionAsc(
                        unitId, fromExclusive, toInclusive, PageRequest.of(0, MAX_OPERATION_TAIL_ENTRIES + 1));
        if (entities.size() > MAX_OPERATION_TAIL_ENTRIES) {
            throw ServiceException.conflict("HISTORY_GAP: replay interval exceeds the bounded operation tail");
        }
        List<OperationRow> rows = entities.stream().map(this::operationRow).toList();
        long bytes = rows.stream().mapToLong(row -> row.envelopeJson().getBytes(StandardCharsets.UTF_8).length).sum();
        if (bytes > MAX_OPERATION_TAIL_BYTES) throw ServiceException.conflict("HISTORY_GAP: replay interval exceeds the bounded operation tail bytes");
        return new OperationTail(unitId, fromExclusive, toInclusive, rows, bytes);
    }

    /**
     * Reads one bounded page for the revision log.  The lower bound is
     * explicit even when the caller requests the newest page.
     */
    public List<OperationRow> readRevisionPage(String unitId, long fromInclusive, long toExclusive, int limit) {
        if (unitId == null || unitId.isBlank()) throw ServiceException.validation("Revision page unitId is required");
        if (fromInclusive < 0 || toExclusive < fromInclusive) throw ServiceException.validation("Revision page interval is invalid");
        int boundedLimit = Math.max(1, Math.min(limit, MAX_OPERATION_TAIL_ENTRIES));
        return operations.findByUnitIdAndRevisionGreaterThanEqualAndRevisionLessThanOrderByRevisionDesc(
                unitId, fromInclusive, toExclusive, PageRequest.of(0, boundedLimit)).stream().map(this::operationRow).toList();
    }

    /**
     * Validates the physical block bytes referenced by a manifest.  This is
     * called only after the caller has acquired the workbook pessimistic lock,
     * so manifest publication and its byte-level proof share one transaction.
     */
    @Transactional(propagation = Propagation.MANDATORY, readOnly = true)
    public void validateDataSourceManifestBlocks(String unitId, ObjectNode source) {
        if (source == null || !source.isObject()) throw ServiceException.validation("Data source manifest is required");
        String sourceId = source.path("id").asText("");
        if (sourceId.isBlank()) throw ServiceException.validation("Data source manifest id is required");
        if (!source.path("blocks").isArray()) throw ServiceException.validation("Data source manifest blocks are required");
        if (source.path("blocks").size() > MAX_MANIFEST_BLOCKS) throw ServiceException.validation("Data source manifest has too many blocks");
        List<ManifestBlock> manifestBlocks = new ArrayList<>();
        Set<String> blockIds = new HashSet<>();
        for (JsonNode raw : source.path("blocks")) {
            if (raw == null || !raw.isObject()) throw ServiceException.validation("Data source block is invalid");
            String blockId = raw.path("id").asText("");
            String expectedChecksum = raw.path("checksum").asText("");
            JsonNode expectedLengthNode = raw.get("byteLength");
            long expectedLength = expectedLengthNode == null ? -1 : expectedLengthNode.asLong(-1);
            if (blockId.isBlank() || !sourceId.equals(raw.path("dataSourceId").asText())
                    || !expectedChecksum.matches("[A-Fa-f0-9]{64}") || expectedLengthNode == null
                    || !expectedLengthNode.isIntegralNumber() || expectedLength < 1) {
                throw ServiceException.validation("Data source block integrity metadata is invalid: " + blockId);
            }
            if (!blockIds.add(blockId)) throw ServiceException.validation("Data source manifest contains a duplicate block: " + blockId);
            manifestBlocks.add(new ManifestBlock(blockId, expectedChecksum, expectedLength));
        }
        if (manifestBlocks.isEmpty()) return;
        Map<String, DataBlockEntity> entities = new HashMap<>();
        for (DataBlockEntity entity : dataBlocks.findManifestBlocks(unitId, sourceId, blockIds)) {
            entities.put(entity.getId().getBlockId(), entity);
        }
        for (ManifestBlock manifestBlock : manifestBlocks) {
            DataBlockEntity entity = entities.get(manifestBlock.blockId());
            if (entity == null) throw ServiceException.conflict("DATA_BLOCK_MISSING: " + sourceId + "/" + manifestBlock.blockId());
            byte[] content = entity.getContent();
            if (content == null || entity.getByteLength() != content.length || manifestBlock.byteLength() != content.length) {
                throw ServiceException.conflict("DATA_BLOCK_LENGTH_MISMATCH: " + sourceId + "/" + manifestBlock.blockId());
            }
            String actualChecksum = sha256(content);
            if (!actualChecksum.equalsIgnoreCase(entity.getChecksum()) || !actualChecksum.equalsIgnoreCase(manifestBlock.checksum())) {
                throw ServiceException.conflict("DATA_BLOCK_CHECKSUM_MISMATCH: " + sourceId + "/" + manifestBlock.blockId());
            }
        }
    }

    @Transactional
    public void insertCheckpoint(String unitId, long revision, String snapshotJson, String checksum, Instant now) {
        CheckpointEntity entity = checkpoints.findAtRevision(unitId, revision)
                .orElseGet(() -> new CheckpointEntity(unitId, revision, snapshotJson, checksum, now));
        entity.update(snapshotJson, checksum, now);
        checkpoints.save(entity);
    }

    public Optional<CheckpointRow> findCheckpoint(String unitId, long revision) {
        return checkpoints.findAtRevision(unitId, revision).map(this::checkpointRow);
    }

    public Optional<CheckpointRow> findLatestCheckpointAtOrBefore(String unitId, long revision) {
        return checkpoints.findLatestAtOrBefore(unitId, revision, PageRequest.of(0, 1)).stream()
                .map(this::checkpointRow).findFirst();
    }

    public void insertAudit(AuditRecord record) {
        String details = record.details() == null ? "{}" : record.details().toString();
        audits.save(new AuditEntity(record.auditId(), record.operationId(), record.unitId(), record.actorId(), record.eventType(),
                record.outcome(), record.reason(), details, record.occurredAt()));
    }

    public List<AuditRecord> listAudit(String unitId, int limit) {
        int boundedLimit = Math.max(1, Math.min(limit, 500));
        return audits.findRecent(unitId, boundedLimit).stream().map(this::auditRecord).toList();
    }

    private WorkbookRow workbookRow(WorkbookEntity entity) {
        return new WorkbookRow(entity.getUnitId(), entity.getName(), entity.getSnapshotJson(), entity.getSnapshotRevision(),
                entity.getRevision(), entity.getLifecycle(), entity.getCreatedAt(), entity.getUpdatedAt());
    }

    private OperationRow operationRow(OperationEntity entity) {
        return new OperationRow(entity.getOperationId(), entity.getUnitId(), entity.getRevision(), entity.getActorSubject(),
                entity.getClientSequence(), entity.getBaseRevision(), entity.getEnvelopeJson(), entity.getCommittedAt());
    }

    private ShareRow shareRow(ShareEntity entity) {
        return new ShareRow(entity.getShareId(), entity.getUnitId(), entity.getTokenHash(), entity.getRole(), entity.getExpiresAt(),
                entity.getRevokedAt(), entity.getCreatedBy(), entity.getCreatedAt());
    }

    private OutboxRow outboxRow(OutboxEntity entity) {
        return new OutboxRow(entity.getEventId(), entity.getUnitId(), entity.getOperationId(), entity.getRevision(), entity.getPayloadJson(),
                entity.getCreatedAt(), entity.getAttempts());
    }

    private CheckpointRow checkpointRow(CheckpointEntity entity) {
        return new CheckpointRow(entity.getId().getUnitId(), entity.getId().getRevision(), entity.getSnapshotJson(), entity.getChecksum(), entity.getCreatedAt());
    }

    private AuditRecord auditRecord(AuditEntity entity) {
        return new AuditRecord(entity.getAuditId(), entity.getOperationId(), entity.getUnitId(), entity.getActorSubject(), entity.getEventType(),
                entity.getOutcome(), entity.getReason(), readJson(entity.getDetailsJson()), entity.getOccurredAt());
    }

    private JsonNode readJson(String value) {
        try {
            return value == null ? mapper.createObjectNode() : mapper.readTree(value);
        } catch (Exception error) {
            throw new IllegalStateException("Stored JSON is invalid", error);
        }
    }

    private static String sha256(byte[] content) {
        try {
            return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content));
        } catch (java.security.NoSuchAlgorithmException error) {
            throw new IllegalStateException("SHA-256 is unavailable", error);
        }
    }

    private record ManifestBlock(String blockId, String checksum, long byteLength) {
    }
}
