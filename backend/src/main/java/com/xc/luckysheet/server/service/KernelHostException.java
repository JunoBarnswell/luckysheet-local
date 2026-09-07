package com.xc.luckysheet.server.service;

/** Typed fail-close error for the Rust kernel process boundary. */
public final class KernelHostException extends ServiceException {
    private final String code;
    private final String object;
    private final String recovery;

    public KernelHostException(String code, String message, String object, String recovery) {
        super(code, statusFor(code), message);
        this.code = code;
        this.object = object;
        this.recovery = recovery == null || recovery.isBlank() ? "retry-after-kernel-recovery" : recovery;
    }

    public String code() { return code; }
    public String object() { return object; }
    public String recovery() { return recovery; }

    private static int statusFor(String code) {
        if ("KERNEL_UNAVAILABLE".equals(code)) return 503;
        if ("CONFLICT".equals(code) || "REVISION_CONFLICT".equals(code)) return 409;
        if ("FORBIDDEN".equals(code)) return 403;
        if ("NOT_FOUND".equals(code)) return 404;
        return 400;
    }
}
