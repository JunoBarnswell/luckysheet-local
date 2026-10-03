package com.xc.luckysheet.server.security;

/** Identity information only. sessionId is not an HTTP cookie or bearer credential. */
public record VerifiedAuthContext(String authority, String subject, String principal, String scopeId,
                                  String sessionId, String tenantId, String appCode, String employmentId,
                                  long contextVersion, String contextId) { }
