package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Index;
import jakarta.persistence.Table;
import jakarta.persistence.Id;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.time.Instant;

@Entity
@Table(name = "operation_log", indexes = @Index(name = "operation_log_unit_revision_idx", columnList = "unit_id,revision"),
        uniqueConstraints = @jakarta.persistence.UniqueConstraint(name = "operation_log_unit_actor_sequence_uk", columnNames = {"unit_id", "actor_subject", "client_sequence"}))
public class OperationEntity {
    @Id
    @Column(name = "operation_id", nullable = false, length = 200)
    private String operationId;

    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;

    @Column(name = "revision", nullable = false)
    private long revision;

    @Column(name = "actor_subject", nullable = false, length = 500)
    private String actorSubject;

    @Column(name = "client_sequence", nullable = false)
    private long clientSequence;

    @Column(name = "base_revision", nullable = false)
    private long baseRevision;

    @JdbcTypeCode(SqlTypes.LONGVARCHAR)
    @Column(name = "envelope_json", nullable = false)
    private String envelopeJson;

    @Column(name = "committed_at", nullable = false)
    private Instant committedAt;

    protected OperationEntity() {
    }

    public OperationEntity(String operationId, String unitId, long revision, String actorSubject, long clientSequence,
                           long baseRevision, String envelopeJson, Instant committedAt) {
        this.operationId = operationId;
        this.unitId = unitId;
        this.revision = revision;
        this.actorSubject = actorSubject;
        this.clientSequence = clientSequence;
        this.baseRevision = baseRevision;
        this.envelopeJson = envelopeJson;
        this.committedAt = committedAt;
    }

    public String getOperationId() { return operationId; }
    public String getUnitId() { return unitId; }
    public long getRevision() { return revision; }
    public String getActorSubject() { return actorSubject; }
    public long getClientSequence() { return clientSequence; }
    public long getBaseRevision() { return baseRevision; }
    public String getEnvelopeJson() { return envelopeJson; }
    public Instant getCommittedAt() { return committedAt; }
}
