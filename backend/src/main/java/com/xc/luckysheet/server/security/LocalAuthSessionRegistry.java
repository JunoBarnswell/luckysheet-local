package com.xc.luckysheet.server.security;

import jakarta.servlet.http.HttpSession;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.web.socket.CloseStatus;
import org.springframework.web.socket.WebSocketSession;

import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/** Tracks local-auth sessions so credential changes take effect immediately. */
@Component
public final class LocalAuthSessionRegistry {
    private static final Logger LOGGER = LoggerFactory.getLogger(LocalAuthSessionRegistry.class);

    private final Map<String, Set<HttpSession>> httpSessions = new ConcurrentHashMap<>();
    private final Map<String, Set<WebSocketSession>> webSocketSessions = new ConcurrentHashMap<>();

    public void registerHttpSession(String subject, HttpSession session) {
        if (subject == null || subject.isBlank() || session == null) return;
        httpSessions.computeIfAbsent(subject, ignored -> ConcurrentHashMap.newKeySet()).add(session);
    }

    public void unregisterHttpSession(String subject, HttpSession session) {
        remove(httpSessions, subject, session);
    }

    public void registerWebSocket(WebSocketSession session) {
        if (session == null || !(session.getPrincipal() instanceof LocalUserAuthentication authentication)) return;
        webSocketSessions.computeIfAbsent(authentication.getName(), ignored -> ConcurrentHashMap.newKeySet()).add(session);
    }

    /** Invalidates HTTP sessions and closes sockets for a changed local account. */
    public void invalidate(String subject) {
        if (subject == null || subject.isBlank()) return;
        Set<HttpSession> sessions = httpSessions.remove(subject);
        if (sessions != null) {
            for (HttpSession session : sessions) {
                try {
                    session.invalidate();
                } catch (IllegalStateException ignored) {
                    // The container already invalidated this session.
                }
            }
        }
        Set<WebSocketSession> sockets = webSocketSessions.remove(subject);
        if (sockets != null) {
            for (WebSocketSession socket : sockets) {
                try {
                    if (socket.isOpen()) socket.close(CloseStatus.POLICY_VIOLATION);
                } catch (IOException error) {
                    LOGGER.debug("Local-auth WebSocket was already closed", error);
                }
            }
        }
    }

    private <T> void remove(Map<String, Set<T>> registry, String subject, T value) {
        if (subject == null || value == null) return;
        Set<T> values = registry.get(subject);
        if (values == null) return;
        values.remove(value);
        if (values.isEmpty()) registry.remove(subject, values);
    }
}
