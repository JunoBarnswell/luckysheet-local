package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookQueryExecutionEntity;
import com.xc.luckysheet.server.persistence.WorkbookQueryExecutionEntityRepository;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Duration;
import java.time.Instant;
import java.util.HexFormat;
import java.util.UUID;

/** Transactional ownership of query execution, publication, cancellation and load. */
@Service
public class QueryExecutionProofService {
    private final WorkbookEntityRepository workbooks;
    private final WorkbookQueryExecutionEntityRepository executions;
    private final AccessControlService access;
    private final WorkbookLifecycleService lifecycle;
    private final ObjectMapper mapper;

    public QueryExecutionProofService(WorkbookEntityRepository workbooks, WorkbookQueryExecutionEntityRepository executions,
            AccessControlService access, WorkbookLifecycleService lifecycle, ObjectMapper mapper) {
        this.workbooks = workbooks;
        this.executions = executions;
        this.access = access;
        this.lifecycle = lifecycle;
        this.mapper = mapper;
    }

    @Transactional
    public StartedExecution begin(String unitId, String queryId, String actor, Duration lifetime) {
        WorkbookEntity workbook = lockWorkbook(unitId, actor);
        return beginLocked(workbook, unitId, queryId, actor, lifetime);
    }

    /**
     * Starts the native analytics proof at the exact workbook revision named
     * by the caller. Analytics is read-only, so viewers may prepare a task;
     * the revision is still checked while the workbook row is locked.
     */
    @Transactional
    public StartedExecution beginAnalytics(String unitId, String queryId, String actor, long revision, Duration lifetime) {
        WorkbookEntity workbook = lockWorkbook(unitId, actor, WorkbookAclRole.VIEWER);
        if (revision != workbook.getRevision()) {
            throw new ServiceException("QUERY_STALE_REVISION", 409,
                    "Analytics must be prepared at the current workbook revision: " + queryId);
        }
        return beginLocked(workbook, unitId, queryId, actor, lifetime);
    }

    /** Authorizes a page request without exposing the persisted result body. */
    @Transactional
    public AnalyticsExecution authorizeAnalytics(String unitId, String queryId, String token, String actor) {
        WorkbookEntity workbook = lockWorkbook(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookQueryExecutionEntity execution = requireExecution(unitId, queryId, token, actor);
        if (!("RUNNING".equals(execution.getStatus()) || "READY".equals(execution.getStatus()))
                || !execution.getExpiresAt().isAfter(Instant.now())) {
            throw new ServiceException("QUERY_EXECUTION_INVALIDATED", 409,
                    "Analytics task was cancelled, consumed or expired: " + queryId);
        }
        if (execution.getSourceRevision() != workbook.getRevision()) {
            throw new ServiceException("QUERY_STALE_REVISION", 409,
                    "Workbook changed after analytics was prepared: " + queryId);
        }
        return new AnalyticsExecution(execution.getExecutionToken(), execution.getSourceRevision(), execution.getExpiresAt());
    }

    /** Seals one bounded native result page. Repeated viewport/drilldown pages
     * update the same proof while the token remains single-owner and unexpired.
     */
    @Transactional
    public String publishAnalytics(String unitId, String queryId, String token, String actor, JsonNode result) {
        WorkbookEntity workbook = lockWorkbook(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookQueryExecutionEntity execution = requireExecution(unitId, queryId, token, actor);
        if (!("RUNNING".equals(execution.getStatus()) || "READY".equals(execution.getStatus()))) {
            throw new ServiceException("QUERY_EXECUTION_INVALIDATED", 409,
                    "Analytics task is no longer publishable: " + queryId);
        }
        if (!execution.getExpiresAt().isAfter(Instant.now()) || execution.getSourceRevision() != workbook.getRevision()) {
            throw new ServiceException("QUERY_STALE_REVISION", 409,
                    "Analytics source revision is stale: " + queryId);
        }
        if (result == null || !result.isObject() || !result.path("revision").isIntegralNumber()
                || result.path("revision").asLong(-1) != execution.getSourceRevision()) {
            throw new ServiceException("QUERY_RESULT_INVALID", 400,
                    "Native analytics result is not pinned to the execution revision: " + queryId);
        }
        try {
            String json = mapper.writeValueAsString(result);
            String hash = checksum(json);
            execution.publish(hash, json);
            executions.save(execution);
            return hash;
        } catch (ServiceException error) { throw error; }
        catch (Exception error) {
            throw new ServiceException("QUERY_RESULT_INVALID", 500,
                    "Unable to seal analytics result: " + queryId, error);
        }
    }

    @Transactional
    public void cancelAnalytics(String unitId, String queryId, String actor) {
        WorkbookEntity workbook = lockWorkbook(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookQueryExecutionEntity execution = executions.findForUpdate(unitId, queryId)
                .orElseThrow(() -> ServiceException.notFound("Analytics task not found: " + queryId));
        WorkbookAclRole role = access.currentRole(unitId, actor);
        if (!execution.getActorSubject().equals(actor) && (role == null || !role.includes(WorkbookAclRole.OWNER))) {
            throw ServiceException.forbidden("Only the analytics actor or workbook owner may cancel a task");
        }
        if ("CONSUMED".equals(execution.getStatus())) {
            throw new ServiceException("QUERY_ALREADY_CONSUMED", 409, "Analytics result was already committed: " + queryId);
        }
        execution.cancel();
        executions.save(execution);
    }

    @Transactional
    public String publish(String unitId, String queryId, String token, String actor, JsonNode result) {
        WorkbookEntity workbook = lockWorkbook(unitId, actor);
        WorkbookQueryExecutionEntity execution = requireExecution(unitId, queryId, token, actor);
        requireCurrent(execution, workbook.getRevision(), "RUNNING");
        try {
            String resultJson = mapper.writeValueAsString(result);
            String hash = checksum(resultJson);
            execution.publish(hash, resultJson);
            executions.save(execution);
            return hash;
        } catch (ServiceException error) { throw error; }
        catch (Exception error) { throw new ServiceException("QUERY_RESULT_INVALID", 500, "Unable to seal query result: " + queryId, error); }
    }

    @Transactional
    public void cancel(String unitId, String queryId, String actor) {
        lockWorkbook(unitId, actor);
        WorkbookQueryExecutionEntity execution = executions.findForUpdate(unitId, queryId)
                .orElseThrow(() -> ServiceException.notFound("Query execution not found: " + queryId));
        WorkbookAclRole role = access.currentRole(unitId, actor);
        if (!execution.getActorSubject().equals(actor) && (role == null || !role.includes(WorkbookAclRole.OWNER))) {
            throw ServiceException.forbidden("Only the query actor or workbook owner may cancel a query");
        }
        if ("CONSUMED".equals(execution.getStatus())) {
            throw new ServiceException("QUERY_ALREADY_CONSUMED", 409, "Query result was already committed: " + queryId);
        }
        execution.cancel();
        executions.save(execution);
    }

    /** Invalidates only this attempt; an older worker cannot cancel a replacement execution. */
    @Transactional
    public void invalidate(String unitId, String queryId, String token) {
        workbooks.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
        executions.findForUpdate(unitId, queryId).filter(row -> row.getExecutionToken().equals(token))
                .filter(row -> !"CONSUMED".equals(row.getStatus())).ifPresent(row -> {
                    row.cancel();
                    executions.save(row);
                });
    }

    /**
     * Must participate in the operation transaction. The caller passes the returned server result to
     * native query-load; client block bytes are never evidence of an executed query.
     */
    @Transactional(propagation = Propagation.MANDATORY)
    public JsonNode consumeProof(String unitId, String queryId, String token, String actor, long sourceRevision, String resultHash) {
        WorkbookEntity workbook = lockWorkbook(unitId, actor);
        WorkbookQueryExecutionEntity execution = requireExecution(unitId, queryId, token, actor);
        requireCurrent(execution, workbook.getRevision(), "READY");
        if (sourceRevision != execution.getSourceRevision() || resultHash == null || !resultHash.equals(execution.getResultHash())) {
            throw new ServiceException("QUERY_RESULT_MISMATCH", 409, "Query revision or result hash does not match execution: " + queryId);
        }
        try {
            if (!checksum(execution.getResultJson()).equals(execution.getResultHash())) {
                throw new IllegalStateException("Stored query result checksum does not match its execution proof");
            }
            JsonNode result = mapper.readTree(execution.getResultJson());
            if (result == null || !result.isObject() || result.path("kind").isTextual()) {
                throw new IllegalStateException("Stored result is not a legacy query-load result");
            }
            execution.consume();
            executions.save(execution);
            return result;
        } catch (Exception error) { throw new ServiceException("QUERY_RESULT_INVALID", 500, "Stored query result is corrupt: " + queryId, error); }
    }

    private WorkbookEntity lockWorkbook(String unitId, String actor) {
        return lockWorkbook(unitId, actor, WorkbookAclRole.EDITOR);
    }

    private WorkbookEntity lockWorkbook(String unitId, String actor, WorkbookAclRole requiredRole) {
        WorkbookEntity workbook = workbooks.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
        access.require(unitId, actor, requiredRole);
        lifecycle.requireActive(unitId);
        return workbook;
    }

    private StartedExecution beginLocked(WorkbookEntity workbook, String unitId, String queryId, String actor, Duration lifetime) {
        executions.findForUpdate(unitId, queryId).ifPresent(previous -> {
            if ("RUNNING".equals(previous.getStatus()) && previous.getExpiresAt().isAfter(Instant.now())) {
                throw new ServiceException("QUERY_ALREADY_RUNNING", 409, "Query is already running: " + queryId);
            }
            executions.delete(previous);
            executions.flush();
        });
        Instant now = Instant.now();
        String token = UUID.randomUUID().toString();
        Instant expiresAt = now.plus(lifetime);
        executions.save(new WorkbookQueryExecutionEntity(token, unitId, queryId, actor, workbook.getRevision(), now, expiresAt));
        return new StartedExecution(token, workbook.getRevision(), expiresAt);
    }

    private String checksum(String result) throws java.security.NoSuchAlgorithmException {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(result.getBytes(StandardCharsets.UTF_8)));
    }

    private WorkbookQueryExecutionEntity requireExecution(String unitId, String queryId, String token, String actor) {
        WorkbookQueryExecutionEntity execution = executions.findForUpdate(unitId, queryId)
                .orElseThrow(() -> new ServiceException("QUERY_EXECUTION_NOT_FOUND", 409, "Execute query before loading: " + queryId));
        if (token == null || !token.equals(execution.getExecutionToken()) || !actor.equals(execution.getActorSubject())) {
            throw new ServiceException("QUERY_EXECUTION_MISMATCH", 409, "Query execution token or actor does not match: " + queryId);
        }
        return execution;
    }

    private void requireCurrent(WorkbookQueryExecutionEntity execution, long currentRevision, String status) {
        if (!status.equals(execution.getStatus()) || !execution.getExpiresAt().isAfter(Instant.now())) {
            throw new ServiceException("QUERY_EXECUTION_INVALIDATED", 409, "Query was cancelled, consumed or expired; execute it again: " + execution.getQueryId());
        }
        if (execution.getSourceRevision() != currentRevision) {
            throw new ServiceException("QUERY_STALE_REVISION", 409, "Workbook changed after query started; execute it again: " + execution.getQueryId());
        }
    }

    public record StartedExecution(String executionToken, long sourceRevision, Instant expiresAt) {
        /** Source compatibility for the legacy connector query tests/callers. */
        public StartedExecution(String executionToken, long sourceRevision) {
            this(executionToken, sourceRevision, null);
        }
    }
    public record AnalyticsExecution(String executionToken, long sourceRevision, Instant expiresAt) {}
}
