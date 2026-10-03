package com.xc.luckysheet.server.security;

import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockHttpSession;
import org.springframework.web.socket.WebSocketSession;
import java.time.Instant;
import java.util.Map;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class LocalAuthSessionLifetimeTest {
    @Test void repeatedRegistrationRetainsRevocationAndLogoutClosesOnlyItsOrigin() throws Exception {
        var registry = new LocalAuthSessionRegistry(mock(com.xc.luckysheet.server.service.GuestShareService.class));
        var principal = mock(LocalUserAuthentication.class);
        when(principal.getName()).thenReturn("local:user");
        var first = new MockHttpSession(); var second = new MockHttpSession();
        registry.registerHttpSession(principal.getName(), first);
        registry.registerHttpSession(principal.getName(), first);
        registry.registerHttpSession(principal.getName(), second);
        var firstSocket = mock(WebSocketSession.class); var secondSocket = mock(WebSocketSession.class);
        for (var socket : java.util.List.of(firstSocket, secondSocket)) { when(socket.getPrincipal()).thenReturn(principal); when(socket.isOpen()).thenReturn(true); registry.registerWebSocket(socket); }
        when(firstSocket.getAttributes()).thenReturn(Map.of(LocalAuthSessionRegistry.HTTP_SESSION_ATTRIBUTE, first));
        when(secondSocket.getAttributes()).thenReturn(Map.of(LocalAuthSessionRegistry.HTTP_SESSION_ATTRIBUTE, second));
        assertTrue(LocalAuthSessionRegistry.isValid(firstSocket));
        first.invalidate();
        verify(firstSocket).close(any());
        verify(secondSocket, never()).close(any());
        registry.invalidate(principal.getName());
        assertTrue(second.isInvalid());
        verify(secondSocket).close(any());
    }
    @Test void expiredJwtIsRejectedAndItsOpenSocketIsClosed() throws Exception {
        var token = org.springframework.security.oauth2.jwt.Jwt.withTokenValue("ephemeral-test-token").header("alg", "none").subject("user").issuedAt(Instant.now().minusSeconds(60)).expiresAt(Instant.now().minusSeconds(1)).build();
        var socket = mock(WebSocketSession.class);
        when(socket.getPrincipal()).thenReturn(new org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken(token));
        when(socket.isOpen()).thenReturn(true);
        assertFalse(LocalAuthSessionRegistry.isValid(socket));
        var registry = new LocalAuthSessionRegistry(mock(com.xc.luckysheet.server.service.GuestShareService.class)); registry.registerWebSocket(socket); registry.closeExpiredSockets();
        verify(socket).close(any());
        registry.closeExpiredSockets(); verify(socket, times(1)).close(any());
    }

    @Test void guestExpiryRevocationAndAtomicAdmissionReleaseIdleSockets() throws Exception {
        var shares = mock(com.xc.luckysheet.server.service.GuestShareService.class);
        var registry = new LocalAuthSessionRegistry(shares);
        var id = java.util.UUID.randomUUID();
        var identity = new com.xc.luckysheet.server.service.GuestShareService.GuestIdentity("guest:" + id, id, "book",
                com.xc.luckysheet.server.contract.WorkbookAclRole.VIEWER, Instant.now().plusSeconds(60));
        when(shares.roleFor("book", identity.subject())).thenReturn(identity.role());
        var admitted = new java.util.ArrayList<WebSocketSession>();
        for (int i = 0; i < LocalAuthSessionRegistry.MAX_SUBJECT_SOCKETS; i++) {
            var socket = guestSocket(identity); assertTrue(registry.registerWebSocket(socket)); admitted.add(socket);
        }
        var rejected = guestSocket(identity); assertFalse(registry.registerWebSocket(rejected)); verify(rejected).close(any());
        registry.unregisterWebSocket(admitted.removeFirst());
        var replacement = guestSocket(identity); assertTrue(registry.registerWebSocket(replacement)); admitted.add(replacement);
        registry.closeExpiredSockets(); for (var socket : admitted) verify(socket, never()).close(any());
        when(shares.roleFor("book", identity.subject())).thenReturn(null);
        registry.closeExpiredSockets(); for (var socket : admitted) verify(socket).close(any());
        var expired = new com.xc.luckysheet.server.service.GuestShareService.GuestIdentity(identity.subject(), id, "book", identity.role(), Instant.now().minusSeconds(1));
        assertFalse(LocalAuthSessionRegistry.isValid(guestSocket(expired)));
        var fresh = guestSocket(identity); assertTrue(registry.registerWebSocket(fresh));
        registry.shareRevoked(new com.xc.luckysheet.server.service.GuestShareService.ShareRevoked(id)); verify(fresh).close(any());
    }
    private WebSocketSession guestSocket(com.xc.luckysheet.server.service.GuestShareService.GuestIdentity identity) {
        var socket = mock(WebSocketSession.class);
        when(socket.getPrincipal()).thenReturn(new GuestShareAuthentication(identity)); when(socket.isOpen()).thenReturn(true);
        return socket;
    }
}
