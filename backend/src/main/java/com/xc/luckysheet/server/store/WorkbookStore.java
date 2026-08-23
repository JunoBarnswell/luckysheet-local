package com.xc.luckysheet.server.store;

import com.xc.luckysheet.server.contract.AclEntry;
import com.xc.luckysheet.server.contract.AuditRecord;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookSummary;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;

import java.sql.ResultSet;
import java.sql.Timestamp;
import java.time.Instant;
import java.util.List;
import java.util.Optional;
import java.util.UUID;

@Repository
public class WorkbookStore {
    private final JdbcTemplate jdbc;
    private final ObjectMapper mapper;

    public WorkbookStore(JdbcTemplate jdbc, ObjectMapper mapper) {
        this.jdbc = jdbc;
        this.mapper = mapper;
    }

    public Optional<WorkbookRow> find(String unitId) {
        return queryOne("""
                SELECT unit_id, name, snapshot_json::text AS snapshot_json, snapshot_revision, revision, created_at, updated_at
                FROM workbooks WHERE unit_id = ?
                """, this::mapWorkbook, unitId);
    }

    public Optional<WorkbookRow> findForUpdate(String unitId) {
        return queryOne("""
                SELECT unit_id, name, snapshot_json::text AS snapshot_json, snapshot_revision, revision, created_at, updated_at
                FROM workbooks WHERE unit_id = ? FOR UPDATE
                """, this::mapWorkbook, unitId);
    }

    public List<WorkbookSummary> listForSubject(String subject) {
        return jdbc.query("""
                SELECT w.unit_id, w.name, w.revision, w.updated_at
                FROM workbooks w JOIN workbook_acl a ON a.unit_id = w.unit_id
                WHERE a.subject = ? ORDER BY w.updated_at DESC
                """, (rs, rowNum) -> new WorkbookSummary(
                rs.getString("unit_id"),
                rs.getString("name"),
                rs.getLong("revision"),
                instant(rs, "updated_at")
        ), subject);
    }

    public void insertWorkbook(String unitId, String name, String snapshotJson, Instant now) {
        jdbc.update("""
                INSERT INTO workbooks(unit_id, name, snapshot_json, snapshot_revision, revision, created_at, updated_at)
                VALUES (?, ?, ?::jsonb, 0, 0, ?, ?)
                """, unitId, name, snapshotJson, Timestamp.from(now), Timestamp.from(now));
    }

    public void updateWorkbook(String unitId, long revision, String snapshotJson, long snapshotRevision, Instant now) {
        jdbc.update("""
                UPDATE workbooks SET revision = ?, snapshot_revision = ?, snapshot_json = ?::jsonb, updated_at = ?
                WHERE unit_id = ?
                """, revision, snapshotRevision, snapshotJson, Timestamp.from(now), unitId);
    }

    public void updateWorkbookRevision(String unitId, long revision, Instant now) {
        jdbc.update("UPDATE workbooks SET revision = ?, updated_at = ? WHERE unit_id = ?", revision, Timestamp.from(now), unitId);
    }

    public void insertAcl(String unitId, String subject, WorkbookAclRole role, Instant now) {
        jdbc.update("""
                INSERT INTO workbook_acl(unit_id, subject, role, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                """, unitId, subject, role.name(), Timestamp.from(now), Timestamp.from(now));
    }

    public List<AclEntry> listAcl(String unitId) {
        return jdbc.query("""
                SELECT unit_id, subject, role, created_at, updated_at
                FROM workbook_acl WHERE unit_id = ? ORDER BY subject
                """, (rs, rowNum) -> new AclEntry(
                rs.getString("unit_id"),
                rs.getString("subject"),
                WorkbookAclRole.valueOf(rs.getString("role")),
                instant(rs, "created_at"),
                instant(rs, "updated_at")
        ), unitId);
    }

    public Optional<WorkbookAclRole> findRole(String unitId, String subject) {
        return queryOne("SELECT role FROM workbook_acl WHERE unit_id = ? AND subject = ?", rs -> WorkbookAclRole.valueOf(get(rs, "role")), unitId, subject);
    }

    public void insertShare(ShareRow share) {
        jdbc.update("""
                INSERT INTO workbook_share(share_id, unit_id, token_hash, role, expires_at, revoked_at, created_by, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, share.shareId(), share.unitId(), share.tokenHash(), share.role().name(), Timestamp.from(share.expiresAt()),
                share.revokedAt() == null ? null : Timestamp.from(share.revokedAt()), share.createdBy(), Timestamp.from(share.createdAt()));
    }

    public Optional<ShareRow> findShare(UUID shareId) {
        return queryOne("""
                SELECT share_id, unit_id, token_hash, role, expires_at, revoked_at, created_by, created_at
                FROM workbook_share WHERE share_id = ?
                """, this::mapShare, shareId);
    }

    public List<ShareRow> listShares(String unitId) {
        return jdbc.query("""
                SELECT share_id, unit_id, token_hash, role, expires_at, revoked_at, created_by, created_at
                FROM workbook_share WHERE unit_id = ? ORDER BY created_at DESC
                """, (rs, rowNum) -> mapShare(rs), unitId);
    }

    public Optional<ShareRow> findActiveShare(String unitId, UUID shareId, Instant now) {
        return queryOne("""
                SELECT share_id, unit_id, token_hash, role, expires_at, revoked_at, created_by, created_at
                FROM workbook_share
                WHERE unit_id = ? AND share_id = ? AND revoked_at IS NULL AND expires_at > ?
                """, this::mapShare, unitId, shareId, Timestamp.from(now));
    }

    public Optional<ShareRow> findActiveShare(UUID shareId, Instant now) {
        return queryOne("""
                SELECT share_id, unit_id, token_hash, role, expires_at, revoked_at, created_by, created_at
                FROM workbook_share
                WHERE share_id = ? AND revoked_at IS NULL AND expires_at > ?
                """, this::mapShare, shareId, Timestamp.from(now));
    }

    public int revokeShare(String unitId, UUID shareId, Instant revokedAt) {
        return jdbc.update("""
                UPDATE workbook_share SET revoked_at = ?
                 WHERE unit_id = ? AND share_id = ? AND revoked_at IS NULL
                """, Timestamp.from(revokedAt), unitId, shareId);
    }

    public void upsertAcl(String unitId, String subject, WorkbookAclRole role, Instant now) {
        jdbc.update("""
                INSERT INTO workbook_acl(unit_id, subject, role, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT (unit_id, subject) DO UPDATE SET role = EXCLUDED.role, updated_at = EXCLUDED.updated_at
                """, unitId, subject, role.name(), Timestamp.from(now), Timestamp.from(now));
    }

    public int deleteAcl(String unitId, String subject) {
        return jdbc.update("DELETE FROM workbook_acl WHERE unit_id = ? AND subject = ?", unitId, subject);
    }

    public Optional<OperationRow> findOperation(String operationId) {
        return queryOne("""
                SELECT operation_id, unit_id, revision, actor_subject, client_sequence, base_revision, envelope_json::text AS envelope_json, committed_at
                FROM operation_log WHERE operation_id = ?
                """, this::mapOperation, operationId);
    }

    public Optional<OperationRow> findOperationBySequence(String unitId, String actorSubject, long sequence) {
        return queryOne("""
                SELECT operation_id, unit_id, revision, actor_subject, client_sequence, base_revision, envelope_json::text AS envelope_json, committed_at
                FROM operation_log WHERE unit_id = ? AND actor_subject = ? AND client_sequence = ?
                """, this::mapOperation, unitId, actorSubject, sequence);
    }

    public void insertOperation(OperationRow operation) {
        jdbc.update("""
                INSERT INTO operation_log(operation_id, unit_id, revision, actor_subject, client_sequence, base_revision, envelope_json, committed_at)
                VALUES (?, ?, ?, ?, ?, ?, ?::jsonb, ?)
                """, operation.operationId(), operation.unitId(), operation.revision(), operation.actorSubject(),
                operation.clientSequence(), operation.baseRevision(), operation.envelopeJson(), Timestamp.from(operation.committedAt()));
    }

    public void insertOutbox(OutboxRow event) {
        jdbc.update("""
                INSERT INTO coordination_outbox(event_id, unit_id, operation_id, revision, payload_json, created_at, next_attempt_at)
                VALUES (?, ?, ?, ?, ?::jsonb, ?, ?)
                ON CONFLICT (unit_id, revision) DO NOTHING
                """, event.eventId(), event.unitId(), event.operationId(), event.revision(), event.payloadJson(),
                Timestamp.from(event.createdAt()), Timestamp.from(event.createdAt()));
    }

    @Transactional
    public List<OutboxRow> claimOutbox(Instant now, Instant leaseUntil, int limit) {
        return jdbc.query("""
                WITH picked AS (
                    SELECT event_id
                    FROM coordination_outbox
                    WHERE published_at IS NULL
                      AND next_attempt_at <= ?
                      AND (lease_until IS NULL OR lease_until < ?)
                    ORDER BY created_at
                    LIMIT ?
                    FOR UPDATE SKIP LOCKED
                )
                UPDATE coordination_outbox outbox
                   SET lease_until = ?, attempts = outbox.attempts + 1
                  FROM picked
                 WHERE outbox.event_id = picked.event_id
                RETURNING outbox.event_id, outbox.unit_id, outbox.operation_id, outbox.revision,
                          outbox.payload_json::text AS payload_json, outbox.created_at, outbox.attempts
                """, (rs, rowNum) -> new OutboxRow(
                UUID.fromString(rs.getString("event_id")),
                rs.getString("unit_id"),
                rs.getString("operation_id"),
                rs.getLong("revision"),
                rs.getString("payload_json"),
                instant(rs, "created_at"),
                rs.getInt("attempts")
        ), Timestamp.from(now), Timestamp.from(now), Math.max(1, Math.min(limit, 1000)), Timestamp.from(leaseUntil));
    }

    public void markOutboxPublished(UUID eventId, Instant publishedAt) {
        jdbc.update("""
                UPDATE coordination_outbox
                   SET published_at = ?, lease_until = NULL
                 WHERE event_id = ? AND published_at IS NULL
                """, Timestamp.from(publishedAt), eventId);
    }

    public void releaseOutbox(UUID eventId, Instant nextAttemptAt) {
        jdbc.update("""
                UPDATE coordination_outbox
                   SET lease_until = NULL, next_attempt_at = ?
                 WHERE event_id = ? AND published_at IS NULL
                """, Timestamp.from(nextAttemptAt), eventId);
    }

    public int deletePublishedOutboxBefore(Instant cutoff) {
        return jdbc.update("""
                DELETE FROM coordination_outbox
                 WHERE published_at IS NOT NULL AND published_at < ?
                """, Timestamp.from(cutoff));
    }

    public List<OperationRow> listOperations(String unitId) {
        return jdbc.query("""
                SELECT operation_id, unit_id, revision, actor_subject, client_sequence, base_revision, envelope_json::text AS envelope_json, committed_at
                FROM operation_log WHERE unit_id = ? ORDER BY revision DESC
                """, (rs, rowNum) -> mapOperation(rs), unitId);
    }

    public void insertCheckpoint(String unitId, long revision, String snapshotJson, String checksum, Instant now) {
        jdbc.update("""
                INSERT INTO snapshot_checkpoint(unit_id, revision, snapshot_json, checksum, created_at)
                VALUES (?, ?, ?::jsonb, ?, ?)
                ON CONFLICT(unit_id, revision) DO UPDATE SET snapshot_json = EXCLUDED.snapshot_json, checksum = EXCLUDED.checksum, created_at = EXCLUDED.created_at
                """, unitId, revision, snapshotJson, checksum, Timestamp.from(now));
    }

    public Optional<CheckpointRow> findCheckpoint(String unitId, long revision) {
        return queryOne("""
                SELECT unit_id, revision, snapshot_json::text AS snapshot_json, checksum, created_at
                FROM snapshot_checkpoint WHERE unit_id = ? AND revision = ?
                """, (rs) -> new CheckpointRow(
                get(rs, "unit_id"), getLong(rs, "revision"), get(rs, "snapshot_json"),
                get(rs, "checksum"), instant(rs, "created_at")
        ), unitId, revision);
    }

    public Optional<CheckpointRow> findLatestCheckpointAtOrBefore(String unitId, long revision) {
        return queryOne("""
                SELECT unit_id, revision, snapshot_json::text AS snapshot_json, checksum, created_at
                FROM snapshot_checkpoint WHERE unit_id = ? AND revision <= ? ORDER BY revision DESC LIMIT 1
                """, (rs) -> new CheckpointRow(
                get(rs, "unit_id"), getLong(rs, "revision"), get(rs, "snapshot_json"),
                get(rs, "checksum"), instant(rs, "created_at")
        ), unitId, revision);
    }

    public void insertAudit(AuditRecord record) {
        String details = record.details() == null ? "{}" : record.details().toString();
        jdbc.update("""
                INSERT INTO operation_audit(audit_id, operation_id, unit_id, actor_subject, event_type, outcome, reason, details_json, occurred_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?::jsonb, ?)
                """, record.auditId(), record.operationId(), record.unitId(), record.actorId(), record.eventType(),
                record.outcome(), record.reason(), details, Timestamp.from(record.occurredAt()));
    }

    public List<AuditRecord> listAudit(String unitId, int limit) {
        return jdbc.query("""
                SELECT audit_id, operation_id, unit_id, actor_subject, event_type, outcome, reason, details_json::text AS details_json, occurred_at
                FROM operation_audit WHERE unit_id = ? ORDER BY occurred_at DESC LIMIT ?
                """, (rs, rowNum) -> new AuditRecord(
                UUID.fromString(rs.getString("audit_id")),
                uuidOrNull(rs.getString("operation_id")),
                uuidOrNull(rs.getString("unit_id")),
                rs.getString("actor_subject"),
                rs.getString("event_type"),
                rs.getString("outcome"),
                rs.getString("reason"),
                readJson(rs.getString("details_json")),
                instant(rs, "occurred_at")
        ), unitId, Math.max(1, Math.min(limit, 500)));
    }

    private JsonNode readJson(String value) {
        try {
            return value == null ? mapper.createObjectNode() : mapper.readTree(value);
        } catch (Exception error) {
            throw new IllegalStateException("Stored JSON is invalid", error);
        }
    }

    private WorkbookRow mapWorkbook(ResultSet rs) {
        return new WorkbookRow(
                get(rs, "unit_id"), get(rs, "name"), get(rs, "snapshot_json"),
                getLong(rs, "snapshot_revision"), getLong(rs, "revision"), instant(rs, "created_at"), instant(rs, "updated_at")
        );
    }

    private OperationRow mapOperation(ResultSet rs) {
        return new OperationRow(
                get(rs, "operation_id"), get(rs, "unit_id"), getLong(rs, "revision"),
                get(rs, "actor_subject"), getLong(rs, "client_sequence"), getLong(rs, "base_revision"), get(rs, "envelope_json"), instant(rs, "committed_at")
        );
    }

    private ShareRow mapShare(ResultSet rs) {
        String revoked = get(rs, "revoked_at");
        return new ShareRow(
                UUID.fromString(get(rs, "share_id")), get(rs, "unit_id"), get(rs, "token_hash"),
                WorkbookAclRole.valueOf(get(rs, "role")), instant(rs, "expires_at"),
                revoked == null ? null : instant(rs, "revoked_at"), get(rs, "created_by"), instant(rs, "created_at")
        );
    }

    private <T> Optional<T> queryOne(String sql, ResultMapper<T> mapper, Object... args) {
        List<T> rows = jdbc.query(sql, (rs, rowNum) -> mapper.map(rs), args);
        return rows.stream().findFirst();
    }

    private static String get(ResultSet rs, String column) {
        try {
            return rs.getString(column);
        } catch (Exception error) {
            throw new IllegalStateException("Unable to read database column " + column, error);
        }
    }

    private static long getLong(ResultSet rs, String column) {
        try {
            return rs.getLong(column);
        } catch (Exception error) {
            throw new IllegalStateException("Unable to read database column " + column, error);
        }
    }

    private static Instant instant(ResultSet rs, String column) {
        try {
            Timestamp timestamp = rs.getTimestamp(column);
            if (timestamp == null) throw new IllegalStateException("Null timestamp in " + column);
            return timestamp.toInstant();
        } catch (Exception error) {
            throw new IllegalStateException("Unable to read timestamp " + column, error);
        }
    }

    private static String uuidOrNull(String value) {
        return value;
    }

    @FunctionalInterface
    private interface ResultMapper<T> {
        T map(ResultSet resultSet);
    }
}
