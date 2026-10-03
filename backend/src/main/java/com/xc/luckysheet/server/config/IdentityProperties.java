package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

/** Claim mapping is trusted deployment configuration, never browser input. */
@ConfigurationProperties(prefix = "luckysheet.identity")
public record IdentityProperties(String authority, String subjectClaim, String sessionClaim,
                                 String tenantClaim, String appClaim, String employmentClaim, String versionClaim) {
    public String subjectClaimName() { return configured(subjectClaim, "sub"); }
    public String sessionClaimName() { return configured(sessionClaim, "sid"); }
    public String tenantClaimName() { return configured(tenantClaim, "tenantId"); }
    public String appClaimName() { return configured(appClaim, "appCode"); }
    public String employmentClaimName() { return configured(employmentClaim, "employmentId"); }
    public String versionClaimName() { return configured(versionClaim, "contextVersion"); }
    private static String configured(String value, String defaultValue) {
        return value == null || value.isBlank() ? defaultValue : value.trim();
    }
}
