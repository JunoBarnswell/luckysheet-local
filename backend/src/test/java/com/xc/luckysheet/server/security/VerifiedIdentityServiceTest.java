package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.config.IdentityProperties;
import com.xc.luckysheet.server.config.AuthProperties;
import com.xc.luckysheet.server.config.SecurityConfig;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.contract.WorkbookSource;
import com.xc.luckysheet.server.contract.WorkbookStorageLocation;
import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.persistence.*;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.WorkbookAuthorizationService;
import com.nimbusds.jose.*;
import com.nimbusds.jose.crypto.RSASSASigner;
import com.nimbusds.jose.jwk.*;
import com.nimbusds.jose.jwk.gen.RSAKeyGenerator;
import com.nimbusds.jwt.*;
import org.junit.jupiter.api.Test;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.security.oauth2.jwt.JwtException;

import java.time.Instant;
import java.util.Date;
import java.util.List;
import java.util.Optional;
import java.net.InetSocketAddress;
import java.nio.charset.StandardCharsets;
import com.sun.net.httpserver.HttpServer;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class VerifiedIdentityServiceTest {
    private final VerifiedIdentityService identities = new VerifiedIdentityService(new IdentityProperties(null, null, null, null, null, null, null));
    private Jwt token(String issuer, String tenant, String employment, String session, long version) {
        var builder = Jwt.withTokenValue("opaque").header("alg", "RS256").issuer(issuer).subject("same-user")
                .claim("sid", session).issuedAt(Instant.now()).expiresAt(Instant.now().plusSeconds(60));
        if (tenant != null) builder.claim("tenantId", tenant).claim("appCode", "erp").claim("contextVersion", version);
        if (employment != null) builder.claim("employmentId", employment);
        return builder.build();
    }

    @Test void identityIsNamespacedButSameContextRenewalPreservesObjectOwnership() {
        var first = identities.jwtContext(token("https://issuer.test", "tenant-a", "job-a", "session-a", 1));
        var renewed = identities.jwtContext(token("https://issuer.test", "tenant-a", "job-a", "session-a", 1));
        assertEquals(first, renewed);
        assertEquals(first.principal(), ActorIdentity.subject(identities.authentication(token("https://issuer.test", "tenant-a", "job-a", "session-a", 1))));
        assertNotEquals(first.principal(), identities.jwtContext(token("https://another.test", "tenant-a", "job-a", "session-a", 1)).principal());
        assertNotEquals(first.scopeId(), identities.jwtContext(token("https://issuer.test", "tenant-b", "job-b", "session-a", 1)).scopeId());
        var switchedEmployment = identities.jwtContext(token("https://issuer.test", "tenant-a", "job-b", "session-a", 2));
        assertEquals(first.principal(), switchedEmployment.principal());
        assertNotEquals(first.contextId(), switchedEmployment.contextId());
        assertNotEquals(first.contextId(), identities.jwtContext(token("https://issuer.test", "tenant-a", "job-a", "session-b", 1)).contextId());
    }

    @Test void missingOrPartialWorkspaceClaimsCannotProduceARegisteredIdentity() {
        Jwt noSession = Jwt.withTokenValue("opaque").header("alg", "RS256").issuer("https://issuer.test").subject("user").expiresAt(Instant.now().plusSeconds(60)).build();
        assertThrows(JwtException.class, () -> identities.jwtContext(noSession));
        assertThrows(JwtException.class, () -> identities.jwtContext(token("https://issuer.test", "tenant", null, "session", 1)));
        Jwt partial = Jwt.withTokenValue("opaque").header("alg", "RS256").issuer("https://issuer.test").subject("user").claim("sid", "session")
                .claim("tenantId", "tenant").claim("appCode", "erp").claim("employmentId", "job").expiresAt(Instant.now().plusSeconds(60)).build();
        assertThrows(JwtException.class, () -> identities.jwtContext(partial));
    }

    @Test void explicitTrustedUserMappingDoesNotAcceptBrowserClaimsAsIdentity() {
        var mapped = new VerifiedIdentityService(new IdentityProperties("erp-authority", "erp_user_id", "sid", null, null, null, null));
        assertThrows(JwtException.class, () -> mapped.jwtContext(token("https://issuer.test", null, null, "session", 0)));
        Jwt jwt = Jwt.withTokenValue("opaque").header("alg", "RS256").issuer("https://issuer.test").subject("sso-user")
                .claim("erp_user_id", "erp-user").claim("sid", "session").expiresAt(Instant.now().plusSeconds(60)).build();
        assertEquals("erp-user", mapped.jwtContext(jwt).subject());
        assertEquals("erp-authority", mapped.jwtContext(jwt).authority());
    }

    @Test void crossWorkspaceAclCannotOverrideTheWorkbookIdentityScope() {
        var first = identities.jwtContext(token("https://issuer.test", "tenant-a", "job", "session", 1));
        var foreign = identities.jwtContext(token("https://issuer.test", "tenant-b", "job", "session", 1));
        WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
        WorkbookAclEntityRepository acl = mock(WorkbookAclEntityRepository.class);
        SpaceMemberEntityRepository members = mock(SpaceMemberEntityRepository.class);
        Instant now = Instant.now();
        when(workbooks.findById("book")).thenReturn(Optional.of(new WorkbookEntity("book", "Book", "{}", 0, 0, now, now,
                first.principal(), null, null, WorkbookStorageLocation.REMOTE, WorkbookSource.NATIVE, WorkbookLifecycle.ACTIVE, null)));
        when(acl.findForSubject("book", foreign.principal())).thenReturn(Optional.of(new WorkbookAclEntity("book", foreign.principal(), WorkbookRole.OWNER, now, now)));
        var service = new WorkbookAuthorizationService(workbooks, acl, members);
        assertEquals(Optional.of(WorkbookRole.OWNER), service.role("book", first.principal()));
        assertTrue(service.role("book", foreign.principal()).isEmpty());
        verify(acl, never()).findForSubject("book", foreign.principal());
    }

    @Test void realSignedJwtValidationRejectsIssuerAudienceExpirySignatureAndMissingSession() throws Exception {
        RSAKey key = new RSAKeyGenerator(2048).keyID("test-key").generate();
        RSAKey foreignKey = new RSAKeyGenerator(2048).keyID("test-key").generate();
        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        byte[] jwks = new JWKSet(key.toPublicJWK()).toString().getBytes(StandardCharsets.UTF_8);
        server.createContext("/jwks", exchange -> { exchange.getResponseHeaders().set("Content-Type", "application/json"); exchange.sendResponseHeaders(200, jwks.length); exchange.getResponseBody().write(jwks); exchange.close(); });
        server.start();
        try {
            var decoder = new SecurityConfig().jwtDecoder(new AuthProperties("oidc", "https://issuer.test", "spreadsheet", "http://127.0.0.1:" + server.getAddress().getPort() + "/jwks", null), identities);
            String valid = signed(key, "https://issuer.test", "spreadsheet", Instant.now().plusSeconds(60), "session");
            assertEquals("same-user", decoder.decode(valid).getSubject());
            for (String rejected : List.of(
                    signed(key, "https://wrong.test", "spreadsheet", Instant.now().plusSeconds(60), "session"),
                    signed(key, "https://issuer.test", "other-api", Instant.now().plusSeconds(60), "session"),
                    signed(key, "https://issuer.test", "spreadsheet", Instant.now().minusSeconds(120), "session"),
                    signed(foreignKey, "https://issuer.test", "spreadsheet", Instant.now().plusSeconds(60), "session"),
                    signed(key, "https://issuer.test", "spreadsheet", Instant.now().plusSeconds(60), null))) {
                assertThrows(JwtException.class, () -> decoder.decode(rejected));
            }
        } finally { server.stop(0); }
    }
    private String signed(RSAKey key, String issuer, String audience, Instant expires, String session) throws Exception {
        var claims = new JWTClaimsSet.Builder().issuer(issuer).subject("same-user").audience(audience).issueTime(Date.from(Instant.now().minusSeconds(180))).expirationTime(Date.from(expires));
        if (session != null) claims.claim("sid", session);
        var jwt = new SignedJWT(new JWSHeader.Builder(JWSAlgorithm.RS256).keyID(key.getKeyID()).build(), claims.build());
        jwt.sign(new RSASSASigner(key)); return jwt.serialize();
    }
}
