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
        var registry = new LocalAuthSessionRegistry();
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
        var registry = new LocalAuthSessionRegistry(); registry.registerWebSocket(socket); registry.closeExpiredSockets();
        verify(socket).close(any());
        registry.closeExpiredSockets(); verify(socket, times(1)).close(any());
    }
}
