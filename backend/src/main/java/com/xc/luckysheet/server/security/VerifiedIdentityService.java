package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.config.IdentityProperties;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpSession;
import org.springframework.security.core.Authentication;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.security.oauth2.jwt.JwtException;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import org.springframework.stereotype.Service;

import java.nio.ByteBuffer;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.HexFormat;
import java.util.List;
import java.util.UUID;

/** One verified identity conversion for REST and WebSocket. */
@Service
public final class VerifiedIdentityService {
    private static final String LOCAL_CONTEXT_NONCE = VerifiedIdentityService.class.getName() + ".nonce";
    private final IdentityProperties properties;

    public VerifiedIdentityService(IdentityProperties properties) { this.properties = properties; }

    public JwtAuthenticationToken authentication(Jwt jwt) {
        VerifiedAuthContext context = jwtContext(jwt);
        return new JwtAuthenticationToken(jwt, List.of(), context.principal());
    }

    public VerifiedAuthContext jwtContext(Jwt jwt) {
        if (jwt.getIssuer() == null || jwt.getExpiresAt() == null) throw new JwtException("Verified issuer and expiry are required");
        String authority = properties.authority() == null || properties.authority().isBlank()
                ? jwt.getIssuer().toString() : properties.authority().trim();
        String subject = required(jwt, properties.subjectClaimName());
        String session = required(jwt, properties.sessionClaimName());
        String tenant = optional(jwt, properties.tenantClaimName());
        String app = optional(jwt, properties.appClaimName());
        String employment = optional(jwt, properties.employmentClaimName());
        if ((tenant == null) != (app == null) || (tenant == null) != (employment == null)) {
            throw new JwtException("Workspace claims must be declared together");
        }
        Object rawVersion = jwt.getClaims().get(properties.versionClaimName());
        long version = 0;
        if (rawVersion != null) {
            if (!(rawVersion instanceof Number number) || number.doubleValue() != number.longValue()
                    || number.longValue() < 0 || number.longValue() > 9007199254740991L) throw new JwtException("Invalid context version");
            version = number.longValue();
        } else if (tenant != null) throw new JwtException("Workspace context version is required");
        String scope = hash(authority, tenant, app);
        String principal = "oidc:" + scope + ":" + hash(authority, subject);
        return context(authority, subject, principal, scope, session, tenant, app, employment, version);
    }

    public VerifiedAuthContext context(Authentication authentication, HttpServletRequest request) {
        if (authentication instanceof JwtAuthenticationToken jwt) return jwtContext(jwt.getToken());
        if (authentication instanceof LocalUserAuthentication local) {
            HttpSession session = request.getSession(true);
            String nonce;
            synchronized (session) {
                nonce = (String) session.getAttribute(LOCAL_CONTEXT_NONCE);
                if (nonce == null) { nonce = UUID.randomUUID().toString(); session.setAttribute(LOCAL_CONTEXT_NONCE, nonce); }
            }
            return context("local", local.getName(), local.getName(), "local", nonce, null, null, null,
                    local.principalData().credentialVersion());
        }
        throw new JwtException("A registered verified identity is required");
    }

    public void rotateLocalContext(HttpSession session) {
        session.setAttribute(LOCAL_CONTEXT_NONCE, UUID.randomUUID().toString());
    }

    public static String scopeForActor(String actor) {
        if (actor != null && actor.startsWith("oidc:")) {
            if (!actor.matches("oidc:[a-f0-9]{64}:[a-f0-9]{64}")) throw new JwtException("Invalid qualified principal");
            return actor.substring(5, 69);
        }
        return "local";
    }

    private VerifiedAuthContext context(String authority, String subject, String principal, String scope, String session,
                                        String tenant, String app, String employment, long version) {
        return new VerifiedAuthContext(authority, subject, principal, scope, session, tenant, app, employment, version,
                hash(authority, subject, scope, session, tenant, app, employment, Long.toString(version)));
    }
    private static String required(Jwt jwt, String claim) {
        String value = optional(jwt, claim);
        if (value == null) throw new JwtException("Required identity claim is missing: " + claim);
        return value;
    }
    private static String optional(Jwt jwt, String claim) {
        Object raw = jwt.getClaims().get(claim);
        if (raw == null) return null;
        if (!(raw instanceof String value) || value.isBlank() || value.length() > 500) throw new JwtException("Invalid identity claim: " + claim);
        return value;
    }
    private static String hash(String... values) {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            for (String value : values) {
                byte[] bytes = value == null ? new byte[0] : value.getBytes(StandardCharsets.UTF_8);
                digest.update(ByteBuffer.allocate(4).putInt(value == null ? -1 : bytes.length).array());
                digest.update(bytes);
            }
            return HexFormat.of().formatHex(digest.digest());
        } catch (NoSuchAlgorithmException error) { throw new IllegalStateException(error); }
    }
}
