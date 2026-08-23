package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.ApiErrorResponse;
import com.xc.luckysheet.server.service.ServiceException;
import org.springframework.dao.DataIntegrityViolationException;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.http.converter.HttpMessageNotReadableException;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

@RestControllerAdvice
public class ApiExceptionHandler {
    @ExceptionHandler(ServiceException.class)
    public ResponseEntity<ApiErrorResponse> handleDomain(ServiceException error) {
        return ResponseEntity.status(error.status()).body(new ApiErrorResponse(error.code(), error.getMessage()));
    }

    @ExceptionHandler({IllegalArgumentException.class, HttpMessageNotReadableException.class})
    public ResponseEntity<ApiErrorResponse> handleValidation(Exception error) {
        return ResponseEntity.badRequest().body(new ApiErrorResponse("VALIDATION_ERROR", safeMessage(error, "Request is invalid")));
    }

    @ExceptionHandler({AccessDeniedException.class})
    public ResponseEntity<ApiErrorResponse> handleDenied(Exception error) {
        return ResponseEntity.status(HttpStatus.FORBIDDEN).body(new ApiErrorResponse("FORBIDDEN", "Access denied"));
    }

    @ExceptionHandler(DataIntegrityViolationException.class)
    public ResponseEntity<ApiErrorResponse> handleConflict(DataIntegrityViolationException error) {
        return ResponseEntity.status(HttpStatus.CONFLICT).body(new ApiErrorResponse("CONFLICT", "The operation conflicts with current workbook state"));
    }

    @ExceptionHandler(Exception.class)
    public ResponseEntity<ApiErrorResponse> handleUnexpected(Exception error) {
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(new ApiErrorResponse("INTERNAL_ERROR", "The request could not be completed"));
    }

    private String safeMessage(Exception error, String fallback) {
        String message = error.getMessage();
        return message == null || message.isBlank() ? fallback : message;
    }
}
