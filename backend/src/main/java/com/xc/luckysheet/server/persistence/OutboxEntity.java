package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;
import jakarta.persistence.UniqueConstraint;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;

import java.time.Instant;
import java.util.UUID;

@Entity
@Table(name = "coordination_outbox", indexes = {
        @Index(name = "coordination_outbox_pending_idx", columnList = "next_attempt_at,created_at"),
        @Index(name = "coordination_outbox_lease_idx", columnList = "lease_until")
}, uniqueConstraints = @UniqueConstraint(name = "coordination_outbox_unit_revision_uk", columnNames = {"unit_id", "revision"}))
public class OutboxEntity {
    @Id
    @JdbcTypeCode(SqlTypes.VARCHAR)
    @Column(name = "event_id", nullable = false)
    private UUID eventId;

    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;

    @Column(name = "operation_id", nullable = false, length = 200)
    private String operationId;

    @Column(name = "revision", nullable = false)
    private long revision;

    @JdbcTypeCode(SqlTypes.LONGVARCHAR)
    @Column(name = "payload_json", nullable = false)
    private String payloadJson;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "published_at")
    private Instant publishedAt;

    @Column(name = "lease_until")
    private Instant leaseUntil;

    @Column(name = "next_attempt_at", nullable = false)
    private Instant nextAttemptAt;

    @Column(name = "attempts", nullable = false)
    private int attempts;

    protected OutboxEntity() {
    }

    public OutboxEntity(UUID eventId, String unitId, String operationId, long revision, String payloadJson,
                        Instant createdAt, Instant nextAttemptAt, int attempts) {
        this.eventId = eventId;
        this.unitId = unitId;
        this.operationId = operationId;
        this.revision = revision;
        this.payloadJson = payloadJson;
        this.createdAt = createdAt;
        this.nextAttemptAt = nextAttemptAt;
        this.attempts = attempts;
    }

    public UUID getEventId() { return eventId; }
    public String getUnitId() { return unitId; }
    public String getOperationId() { return operationId; }
    public long getRevision() { return revision; }
    public String getPayloadJson() { return payloadJson; }
    public Instant getCreatedAt() { return createdAt; }
    public Instant getPublishedAt() { return publishedAt; }
    public Instant getLeaseUntil() { return leaseUntil; }
    public Instant getNextAttemptAt() { return nextAttemptAt; }
    public int getAttempts() { return attempts; }

    public void claim(Instant leaseUntil) {
        this.leaseUntil = leaseUntil;
        this.attempts++;
    }

    public void markPublished(Instant publishedAt) {
        this.publishedAt = publishedAt;
        this.leaseUntil = null;
    }

    public void release(Instant nextAttemptAt) {
        this.leaseUntil = null;
        this.nextAttemptAt = nextAttemptAt;
    }
}
