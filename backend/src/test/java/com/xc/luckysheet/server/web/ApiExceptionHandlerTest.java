package com.xc.luckysheet.server.web;

import org.junit.jupiter.api.Test;
import org.springframework.dao.OptimisticLockingFailureException;
import org.springframework.http.HttpStatus;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

class ApiExceptionHandlerTest {
    private final ApiExceptionHandler handler = new ApiExceptionHandler();

    @Test
    void optimisticEntityConflictIsAnObservableRetryableWorkbookConflict() {
        var response = handler.handleOptimisticConflict(new OptimisticLockingFailureException("stale workbook"));

        assertEquals(HttpStatus.CONFLICT, response.getStatusCode());
        assertNotNull(response.getBody());
        assertEquals("CONFLICT", response.getBody().code());
        assertEquals(
                "Workbook changed while this operation was being saved; the draft was retained and must be retried against the latest revision",
                response.getBody().message()
        );
    }

    @Test
    void unrelatedUnexpectedFailureRemainsInternalError() {
        var response = handler.handleUnexpected(new IllegalStateException("unrelated"));

        assertEquals(HttpStatus.INTERNAL_SERVER_ERROR, response.getStatusCode());
        assertNotNull(response.getBody());
        assertEquals("INTERNAL_ERROR", response.getBody().code());
    }
}
