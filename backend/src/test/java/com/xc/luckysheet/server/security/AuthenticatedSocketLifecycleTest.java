package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.service.GuestShareService;
import org.junit.jupiter.api.Test;
import org.springframework.web.socket.WebSocketHandler;
import org.springframework.web.socket.WebSocketSession;
import org.springframework.web.socket.TextMessage;
import org.springframework.web.socket.CloseStatus;

import java.time.Instant;
import java.util.UUID;

import static org.mockito.Mockito.*;
import static org.junit.jupiter.api.Assertions.*;

class AuthenticatedSocketLifecycleTest {
    private WebSocketSession guest() {
        var session = mock(WebSocketSession.class);
        var id = UUID.randomUUID();
        when(session.getPrincipal()).thenReturn(new GuestShareAuthentication(new GuestShareService.GuestIdentity(
                "guest:" + id, id, "book", WorkbookRole.VIEWER, Instant.now().plusSeconds(60))));
        when(session.getId()).thenReturn(id.toString());
        when(session.isOpen()).thenReturn(true);
        return session;
    }

    @Test void acceptedSocketTouchesActivityAndReleasesAdmissionOnClose() throws Exception {
        var registry = mock(LocalAuthSessionRegistry.class);
        var session = guest();
        when(registry.registerWebSocket(session)).thenReturn(true);
        var handler = mock(WebSocketHandler.class);
        var lifecycle = new AuthenticatedSocketLifecycle(registry);
        try {
            var decorated = lifecycle.decorate(handler);
            decorated.afterConnectionEstablished(session);
            verify(handler).afterConnectionEstablished(session);
            var message = new TextMessage("activity");
            decorated.handleMessage(session, message);
            verify(registry).touch(session);
            verify(handler).handleMessage(session, message);
            decorated.afterConnectionClosed(session, CloseStatus.NORMAL);
            verify(registry).unregisterWebSocket(session);
        } finally { lifecycle.close(); }
    }

    @Test void rejectedAdmissionNeverReachesTheWorkbookHandler() throws Exception {
        var registry = mock(LocalAuthSessionRegistry.class);
        var session = guest();
        when(registry.registerWebSocket(session)).thenReturn(false);
        var handler = mock(WebSocketHandler.class);
        var lifecycle = new AuthenticatedSocketLifecycle(registry);
        try {
            lifecycle.decorate(handler).afterConnectionEstablished(session);
            verifyNoInteractions(handler);
        } finally { lifecycle.close(); }
    }

    @Test void failedEstablishmentReleasesAdmissionAndPropagatesTheFailure() throws Exception {
        var registry = mock(LocalAuthSessionRegistry.class);
        var session = guest();
        when(registry.registerWebSocket(session)).thenReturn(true);
        var handler = mock(WebSocketHandler.class);
        var failure = new IllegalStateException("cannot establish");
        doThrow(failure).when(handler).afterConnectionEstablished(session);
        var lifecycle = new AuthenticatedSocketLifecycle(registry);
        try {
            assertSame(failure, assertThrows(IllegalStateException.class,
                    () -> lifecycle.decorate(handler).afterConnectionEstablished(session)));
            verify(registry).unregisterWebSocket(session);
        } finally { lifecycle.close(); }
    }
}
