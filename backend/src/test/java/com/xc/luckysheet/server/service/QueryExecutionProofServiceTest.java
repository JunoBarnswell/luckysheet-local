package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookQueryExecutionEntity;
import com.xc.luckysheet.server.persistence.WorkbookQueryExecutionEntityRepository;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import org.junit.jupiter.api.Test;
import java.time.Instant;
import java.util.Optional;
import java.time.Duration;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class QueryExecutionProofServiceTest {
    private final ObjectMapper mapper = new ObjectMapper();
    private final WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
    private final WorkbookQueryExecutionEntityRepository executions = mock(WorkbookQueryExecutionEntityRepository.class);
    private final AccessControlService access = mock(AccessControlService.class);
    private final QueryExecutionProofService service = new QueryExecutionProofService(workbooks, executions,
            access, mock(WorkbookLifecycleService.class), mapper);

    private WorkbookQueryExecutionEntity execution(long workbookRevision) {
        Instant now = Instant.now();
        when(workbooks.findForUpdate("unit")).thenReturn(Optional.of(new WorkbookEntity("unit", "Workbook", workbookRevision, now, now, "actor", null, null, com.xc.luckysheet.server.contract.WorkbookStorageLocation.REMOTE, com.xc.luckysheet.server.contract.WorkbookSource.NATIVE, com.xc.luckysheet.server.contract.WorkbookLifecycle.ACTIVE, null)));
        WorkbookQueryExecutionEntity execution = new WorkbookQueryExecutionEntity("token", "unit", "query", "actor", 4, now, now.plusSeconds(60));
        when(executions.findForUpdate("unit", "query")).thenReturn(Optional.of(execution));
        return execution;
    }

    @Test void publishedServerResultCanBeLoadedExactlyOnce() throws Exception {
        WorkbookQueryExecutionEntity execution = execution(4);
        var result = mapper.readTree("{\"columns\":[\"amount\"],\"rows\":[[3]]}");
        String hash = service.publish("unit", "query", "token", "actor", result);
        assertEquals(64, hash.length());
        assertEquals(result, service.consumeProof("unit", "query", "token", "actor", 4, hash));
        assertEquals("CONSUMED", execution.getStatus());
        assertNull(execution.getResultJson());
        assertEquals("QUERY_EXECUTION_INVALIDATED", assertThrows(ServiceException.class,
                () -> service.consumeProof("unit", "query", "token", "actor", 4, hash)).code());
    }

    @Test void changedWorkbookCannotPublishResultFromOldRevision() {
        WorkbookQueryExecutionEntity execution = execution(5);
        assertEquals("QUERY_STALE_REVISION", assertThrows(ServiceException.class,
                () -> service.publish("unit", "query", "token", "actor", mapper.createObjectNode())).code());
        assertEquals("RUNNING", execution.getStatus());
        assertNull(execution.getResultJson());
    }

    @Test void cancellationRejectsPublicationAndLoad() {
        WorkbookQueryExecutionEntity execution = execution(4);
        service.cancel("unit", "query", "actor");
        assertEquals("CANCELLED", execution.getStatus());
        assertEquals("QUERY_EXECUTION_INVALIDATED", assertThrows(ServiceException.class,
                () -> service.publish("unit", "query", "token", "actor", mapper.createObjectNode())).code());
        assertEquals("QUERY_EXECUTION_INVALIDATED", assertThrows(ServiceException.class,
                () -> service.consumeProof("unit", "query", "token", "actor", 4, "hash")).code());
    }

    @Test void cancellingReadyResultPreventsLaterCommit() {
        execution(4);
        String hash = service.publish("unit", "query", "token", "actor", mapper.createObjectNode());
        service.cancel("unit", "query", "actor");
        assertEquals("QUERY_EXECUTION_INVALIDATED", assertThrows(ServiceException.class,
                () -> service.consumeProof("unit", "query", "token", "actor", 4, hash)).code());
    }

    @Test void wrongActorTokenAndHashCannotConsumeTheResult() {
        WorkbookQueryExecutionEntity execution = execution(4);
        String hash = service.publish("unit", "query", "token", "actor", mapper.createObjectNode());
        assertEquals("QUERY_EXECUTION_MISMATCH", assertThrows(ServiceException.class,
                () -> service.consumeProof("unit", "query", "other-token", "actor", 4, hash)).code());
        assertEquals("QUERY_EXECUTION_MISMATCH", assertThrows(ServiceException.class,
                () -> service.consumeProof("unit", "query", "token", "other-actor", 4, hash)).code());
        assertEquals("QUERY_RESULT_MISMATCH", assertThrows(ServiceException.class,
                () -> service.consumeProof("unit", "query", "token", "actor", 4, "other-hash")).code());
        assertEquals("READY", execution.getStatus());
    }

    @Test void analyticsProofPinsRevisionAndSealsOnlyMatchingNativePages() throws Exception {
        WorkbookQueryExecutionEntity execution = execution(4);
        when(executions.findForUpdate("unit", "query")).thenReturn(Optional.empty(), Optional.of(execution));
        var started = service.beginAnalytics("unit", "query", "actor", 4, Duration.ofMinutes(5));
        assertEquals(4, started.sourceRevision());
        assertEquals("RUNNING", execution.getStatus());
        var page = mapper.readTree("{\"kind\":\"query\",\"revision\":4,\"columns\":[0],\"rows\":[],\"total\":0,\"grouped\":false}");
        String hash = service.publishAnalytics("unit", "query", started.executionToken(), "actor", page);
        assertEquals(64, hash.length());
        assertEquals("READY", execution.getStatus());
        assertEquals("QUERY_RESULT_INVALID", assertThrows(ServiceException.class,
                () -> service.publishAnalytics("unit", "query", started.executionToken(), "actor",
                        mapper.readTree("{\"kind\":\"query\",\"revision\":3}"))).code());
    }

    @Test void analyticsPrepareRejectsStaleRevisionAndCancellationRejectsAnotherViewer() {
        execution(5);
        assertEquals("QUERY_STALE_REVISION", assertThrows(ServiceException.class,
                () -> service.beginAnalytics("unit", "query", "actor", 4, Duration.ofMinutes(5))).code());
        WorkbookQueryExecutionEntity fresh = execution(4);
        when(executions.findForUpdate("unit", "query")).thenReturn(Optional.empty(), Optional.of(fresh));
        service.beginAnalytics("unit", "query", "actor", 4, Duration.ofMinutes(5));
        when(executions.findForUpdate("unit", "query")).thenReturn(Optional.of(fresh));
        when(access.currentRole("unit", "other")).thenReturn(WorkbookAclRole.VIEWER);
        assertEquals("FORBIDDEN", assertThrows(ServiceException.class,
                () -> service.cancelAnalytics("unit", "query", "other")).code());
    }
}
