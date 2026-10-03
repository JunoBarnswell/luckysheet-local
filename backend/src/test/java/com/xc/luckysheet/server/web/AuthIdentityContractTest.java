package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.config.AuthProperties;
import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.persistence.LocalUserEntity;
import com.xc.luckysheet.server.security.GuestShareAuthentication;
import com.xc.luckysheet.server.security.LocalAuthSessionRegistry;
import com.xc.luckysheet.server.security.LocalUserAuthentication;
import com.xc.luckysheet.server.security.VerifiedAuthContext;
import com.xc.luckysheet.server.security.VerifiedIdentityService;
import com.xc.luckysheet.server.service.GuestShareService;
import com.xc.luckysheet.server.service.LocalAuthService;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockHttpServletRequest;
import org.springframework.mock.web.MockHttpServletResponse;
import org.springframework.security.web.context.SecurityContextRepository;
import org.springframework.security.web.csrf.CsrfToken;
import org.springframework.security.web.csrf.CsrfTokenRepository;
import org.springframework.security.web.csrf.DefaultCsrfToken;

import java.time.Instant;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class AuthIdentityContractTest {
    private AuthController controller(VerifiedIdentityService identities) {
        return new AuthController(new AuthProperties(null, null, null), mock(LocalAuthService.class),
                mock(SecurityContextRepository.class), mock(CsrfTokenRepository.class),
                mock(LocalAuthSessionRegistry.class), identities);
    }
    private MockHttpServletRequest request() {
        var request = new MockHttpServletRequest("GET", "/api/auth/session");
        request.setAttribute(CsrfToken.class.getName(), new DefaultCsrfToken("X-CSRF-TOKEN", "_csrf", "test-csrf"));
        return request;
    }

    @Test
    void workbookCapabilityCannotPublishARegisteredIdentityOrAdministrativeAuthority() {
        var identities = mock(VerifiedIdentityService.class);
        var guest = new GuestShareAuthentication(new GuestShareService.GuestIdentity(
                "guest:test", UUID.randomUUID(), "book", WorkbookRole.VIEWER, Instant.now().plusSeconds(60)));
        var snapshot = controller(identities).session(guest, request(), new MockHttpServletResponse());
        assertFalse(snapshot.authenticated()); assertFalse(snapshot.admin());
        assertNull(snapshot.subject()); assertNull(snapshot.context()); assertNull(snapshot.displayName());
        assertEquals("test-csrf", snapshot.csrfToken());
        verifyNoInteractions(identities);
    }

    @Test
    void registeredIdentityKeepsTheVerifiedContextAndAdministrativeCapability() {
        var identities = mock(VerifiedIdentityService.class);
        var now = Instant.now();
        var authentication = LocalUserAuthentication.from(new LocalUserEntity(
                "admin", "admin", "unused-private-hash", "Admin", true, true, now, now));
        var request = request();
        var context = new VerifiedAuthContext("local", "local:admin", "local:admin", "local", "public-nonce",
                null, null, null, 0, "verified-context");
        when(identities.context(authentication, request)).thenReturn(context);
        var snapshot = controller(identities).session(authentication, request, new MockHttpServletResponse());
        assertTrue(snapshot.authenticated()); assertTrue(snapshot.admin());
        assertEquals("local:admin", snapshot.subject()); assertEquals(context, snapshot.context());
        assertEquals("Admin", snapshot.displayName());
    }
}
