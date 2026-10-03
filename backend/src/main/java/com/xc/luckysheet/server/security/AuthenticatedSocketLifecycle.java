package com.xc.luckysheet.server.security;

import jakarta.annotation.PreDestroy;
import org.springframework.stereotype.Component;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import org.springframework.web.socket.CloseStatus;
import org.springframework.web.socket.WebSocketHandler;
import org.springframework.web.socket.WebSocketSession;
import org.springframework.web.socket.handler.WebSocketHandlerDecorator;

import java.io.IOException;
import java.time.Duration;
import java.time.Instant;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

/** Retires live connections at credential expiry and cleans up local revocation ownership. */
@Component
public final class AuthenticatedSocketLifecycle {
    private final LocalAuthSessionRegistry localSessions;
    private final java.util.concurrent.ScheduledExecutorService scheduler = Executors.newSingleThreadScheduledExecutor(task -> {
        Thread thread = new Thread(task, "spreadsheet-socket-auth-expiry"); thread.setDaemon(true); return thread;
    });
    private final Map<String, ScheduledFuture<?>> expiryTasks = new ConcurrentHashMap<>();
    public AuthenticatedSocketLifecycle(LocalAuthSessionRegistry localSessions) { this.localSessions = localSessions; }

    public WebSocketHandler decorate(WebSocketHandler handler) {
        return new WebSocketHandlerDecorator(handler) {
            @Override
            public void afterConnectionEstablished(WebSocketSession session) throws Exception {
                if (!LocalAuthSessionRegistry.isValid(session)) { session.close(CloseStatus.POLICY_VIOLATION); return; }
                if (!localSessions.registerWebSocket(session)) return;
                Instant expires = session.getPrincipal() instanceof JwtAuthenticationToken jwt ? jwt.getToken().getExpiresAt()
                        : session.getPrincipal() instanceof GuestShareAuthentication guest ? guest.identity().expiresAt() : null;
                if (expires != null) {
                    if (!expires.isAfter(Instant.now())) { session.close(CloseStatus.POLICY_VIOLATION); return; }
                    expiryTasks.put(session.getId(), scheduler.schedule(() -> {
                        try { session.close(CloseStatus.POLICY_VIOLATION); }
                        catch (IOException error) { org.slf4j.LoggerFactory.getLogger(AuthenticatedSocketLifecycle.class).debug("Expired socket already closed", error); }
                    }, Math.max(0, Duration.between(Instant.now(), expires).toMillis()), TimeUnit.MILLISECONDS));
                }
                try { super.afterConnectionEstablished(session); }
                catch (Exception error) {
                    ScheduledFuture<?> task = expiryTasks.remove(session.getId());
                    if (task != null) task.cancel(false);
                    localSessions.unregisterWebSocket(session);
                    throw error;
                }
            }
            @Override
            public void handleMessage(WebSocketSession session, org.springframework.web.socket.WebSocketMessage<?> message) throws Exception {
                localSessions.touch(session);
                super.handleMessage(session, message);
            }
            @Override
            public void afterConnectionClosed(WebSocketSession session, CloseStatus status) throws Exception {
                ScheduledFuture<?> task = expiryTasks.remove(session.getId());
                if (task != null) task.cancel(false);
                localSessions.unregisterWebSocket(session);
                super.afterConnectionClosed(session, status);
            }
        };
    }
    @PreDestroy
    public void close() { scheduler.shutdownNow(); expiryTasks.clear(); }
}
