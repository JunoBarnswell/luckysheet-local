package com.xc.luckysheet.server.security;

import jakarta.servlet.http.HttpSession;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.web.socket.CloseStatus;
import org.springframework.web.socket.WebSocketSession;

import java.io.IOException;
import java.time.Instant;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/** Tracks local-auth sessions so credential changes take effect immediately. */
@Component
public final class LocalAuthSessionRegistry {
    private static final Logger LOGGER = LoggerFactory.getLogger(LocalAuthSessionRegistry.class);

    public static final String HTTP_SESSION_ATTRIBUTE = "localAuthHttpSession";
    private final Set<WebSocketSession> openedSockets = ConcurrentHashMap.newKeySet();

    private final Map<String, Set<HttpSession>> httpSessions = new ConcurrentHashMap<>();
    private final Map<String, Set<WebSocketSession>> webSocketSessions = new ConcurrentHashMap<>();

    public void registerHttpSession(String subject, HttpSession session) {
        if (subject == null || subject.isBlank() || session == null) return;
        session.setAttribute("localAuthSocketLifetime", new jakarta.servlet.http.HttpSessionBindingListener() {
            @Override public void valueUnbound(jakarta.servlet.http.HttpSessionBindingEvent event) {
                closeSessionSockets(session);
                remove(httpSessions, subject, session);
            }
        });
        httpSessions.compute(subject, (key, existing) -> {
            Set<HttpSession> sessions = existing == null ? ConcurrentHashMap.newKeySet() : existing;
            sessions.add(session);
            return sessions;
        });
    }

    public void unregisterHttpSession(String subject, HttpSession session) {
        closeSessionSockets(session);
        remove(httpSessions, subject, session);
    }

    public void registerWebSocket(WebSocketSession session) {
        if (session == null) return;
        openedSockets.add(session);
        if (!(session.getPrincipal() instanceof LocalUserAuthentication authentication)) return;
        webSocketSessions.compute(authentication.getName(), (key, existing) -> {
            Set<WebSocketSession> sockets = existing == null ? ConcurrentHashMap.newKeySet() : existing;
            sockets.add(session);
            return sockets;
        });
    }

    public void unregisterWebSocket(WebSocketSession session) {
        openedSockets.remove(session);
        if (session.getPrincipal() instanceof LocalUserAuthentication authentication) remove(webSocketSessions, authentication.getName(), session);
    }

    public static boolean isValid(WebSocketSession session) {
        if (session.getPrincipal() instanceof org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken jwt) {
            Instant expiration = jwt.getToken().getExpiresAt();
            return expiration != null && expiration.isAfter(Instant.now());
        }
        if (session.getPrincipal() instanceof LocalUserAuthentication) {
            Object origin = session.getAttributes().get(HTTP_SESSION_ATTRIBUTE);
            if (!(origin instanceof HttpSession http)) return false;
            try { return http.getAttribute("localAuthSocketLifetime") != null; }
            catch (IllegalStateException expired) { return false; }
        }
        return session.getPrincipal() instanceof GuestShareAuthentication;
    }

    @org.springframework.scheduling.annotation.Scheduled(fixedDelay = 1000)
    public void closeExpiredSockets() {
        for (WebSocketSession socket : openedSockets) {
            if (!socket.isOpen() || !isValid(socket)) { close(socket); unregisterWebSocket(socket); }
        }
    }

    private void closeSessionSockets(HttpSession session) {
        if (session == null) return;
        for (WebSocketSession socket : openedSockets) {
            if (socket.getAttributes().get(HTTP_SESSION_ATTRIBUTE) == session) { close(socket); unregisterWebSocket(socket); }
        }
    }

    private void close(WebSocketSession socket) {
        try { if (socket.isOpen()) socket.close(CloseStatus.POLICY_VIOLATION); }
        catch (IOException error) { LOGGER.debug("Authentication socket already closed", error); }
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
        registry.computeIfPresent(subject, (key, values) -> {
            values.remove(value);
            return values.isEmpty() ? null : values;
        });
    }
}
