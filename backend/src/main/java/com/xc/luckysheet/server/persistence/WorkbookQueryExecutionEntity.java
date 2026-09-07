package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import jakarta.persistence.UniqueConstraint;
import org.hibernate.annotations.JdbcTypeCode;
import org.hibernate.type.SqlTypes;
import java.time.Instant;

/** A server-owned, single-use result proof, never a client supplied result. */
@Entity
@Table(name = "workbook_query_execution", uniqueConstraints = @UniqueConstraint(
        name = "workbook_query_execution_query_uq", columnNames = {"unit_id", "query_id"}))
public class WorkbookQueryExecutionEntity {
    @Id @Column(name = "execution_token", nullable = false, length = 36)
    private String executionToken;
    @Column(name = "unit_id", nullable = false, length = 200)
    private String unitId;
    @Column(name = "query_id", nullable = false, length = 200)
    private String queryId;
    @Column(name = "actor_subject", nullable = false, length = 500)
    private String actorSubject;
    @Column(name = "source_revision", nullable = false)
    private long sourceRevision;
    @Column(name = "status", nullable = false, length = 16)
    private String status;
    @Column(name = "result_hash", length = 64)
    private String resultHash;
    @JdbcTypeCode(SqlTypes.LONGVARCHAR) @Column(name = "result_json")
    private String resultJson;
    @Column(name = "created_at", nullable = false)
    private Instant createdAt;
    @Column(name = "expires_at", nullable = false)
    private Instant expiresAt;

    protected WorkbookQueryExecutionEntity() {}

    public WorkbookQueryExecutionEntity(String token, String unitId, String queryId, String actor,
            long revision, Instant createdAt, Instant expiresAt) {
        this.executionToken = token;
        this.unitId = unitId;
        this.queryId = queryId;
        this.actorSubject = actor;
        this.sourceRevision = revision;
        this.status = "RUNNING";
        this.createdAt = createdAt;
        this.expiresAt = expiresAt;
    }

    public String getExecutionToken() { return executionToken; }
    public String getUnitId() { return unitId; }
    public String getQueryId() { return queryId; }
    public String getActorSubject() { return actorSubject; }
    public long getSourceRevision() { return sourceRevision; }
    public String getStatus() { return status; }
    public String getResultHash() { return resultHash; }
    public String getResultJson() { return resultJson; }
    public Instant getExpiresAt() { return expiresAt; }
    public void publish(String hash, String result) { status = "READY"; resultHash = hash; resultJson = result; }
    public void cancel() { status = "CANCELLED"; resultHash = null; resultJson = null; }
    public void consume() { status = "CONSUMED"; resultJson = null; }
}
