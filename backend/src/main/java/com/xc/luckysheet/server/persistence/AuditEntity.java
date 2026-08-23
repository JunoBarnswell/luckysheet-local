package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.time.Instant;
import java.util.UUID;

@Entity
@Table(name = "operation_audit", indexes = @Index(name = "operation_audit_unit_time_idx", columnList = "unit_id,occurred_at"))
public class AuditEntity {
    @Id
    @JdbcTypeCode(SqlTypes.VARCHAR)
    @Column(name = "audit_id", nullable = false)
    private UUID auditId;

    @Column(name = "operation_id", length = 200)
    private String operationId;

    @Column(name = "unit_id", length = 200)
    private String unitId;

    @Column(name = "actor_subject", nullable = false, length = 500)
    private String actorSubject;

    @Column(name = "event_type", nullable = false, length = 100)
    private String eventType;

    @Column(name = "outcome", nullable = false, length = 16)
    private String outcome;

    @Column(name = "reason", length = 2000)
    private String reason;

    @JdbcTypeCode(SqlTypes.LONGVARCHAR)
    @Column(name = "details_json", nullable = false)
    private String detailsJson;

    @Column(name = "occurred_at", nullable = false)
    private Instant occurredAt;

    protected AuditEntity() {
    }

    public AuditEntity(UUID auditId, String operationId, String unitId, String actorSubject, String eventType,
                       String outcome, String reason, String detailsJson, Instant occurredAt) {
        this.auditId = auditId;
        this.operationId = operationId;
        this.unitId = unitId;
        this.actorSubject = actorSubject;
        this.eventType = eventType;
        this.outcome = outcome;
        this.reason = reason;
        this.detailsJson = detailsJson;
        this.occurredAt = occurredAt;
    }

    public UUID getAuditId() { return auditId; }
    public String getOperationId() { return operationId; }
    public String getUnitId() { return unitId; }
    public String getActorSubject() { return actorSubject; }
    public String getEventType() { return eventType; }
    public String getOutcome() { return outcome; }
    public String getReason() { return reason; }
    public String getDetailsJson() { return detailsJson; }
    public Instant getOccurredAt() { return occurredAt; }
}
