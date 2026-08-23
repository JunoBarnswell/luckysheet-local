package com.xc.luckysheet.server.service;

public class ServiceException extends RuntimeException {
    private final String code;
    private final int status;

    public ServiceException(String code, int status, String message) {
        super(message);
        this.code = code;
        this.status = status;
    }

    public ServiceException(String code, int status, String message, Throwable cause) {
        super(message, cause);
        this.code = code;
        this.status = status;
    }

    public String code() {
        return code;
    }

    public int status() {
        return status;
    }

    public static ServiceException notFound(String message) {
        return new ServiceException("NOT_FOUND", 404, message);
    }

    public static ServiceException validation(String message) {
        return new ServiceException("VALIDATION_ERROR", 400, message);
    }

    public static ServiceException conflict(String message) {
        return new ServiceException("CONFLICT", 409, message);
    }

    public static ServiceException trashed(String message) {
        return new ServiceException("WORKBOOK_TRASHED", 409, message);
    }

    public static ServiceException forbidden(String message) {
        return new ServiceException("FORBIDDEN", 403, message);
    }

    public static ServiceException timeout(String message) {
        return new ServiceException("TIMEOUT", 408, message);
    }

    public static ServiceException unavailable(String message) {
        return new ServiceException("SERVICE_UNAVAILABLE", 503, message);
    }

    public static ServiceException unauthenticated(String message) {
        return new ServiceException("UNAUTHENTICATED", 401, message);
    }
}
