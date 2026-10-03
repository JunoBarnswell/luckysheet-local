package com.xc.luckysheet.server.security;

import org.springframework.web.socket.WebSocketHandler;
import org.springframework.web.socket.handler.WebSocketHandlerDecorator;
import org.springframework.web.socket.handler.WebSocketHandlerDecoratorFactory;

/** Adds the opened socket to the local-auth revocation registry. */
public final class LocalAuthWebSocketHandlerDecoratorFactory implements WebSocketHandlerDecoratorFactory {
    private final LocalAuthSessionRegistry sessions;

    public LocalAuthWebSocketHandlerDecoratorFactory(LocalAuthSessionRegistry sessions) {
        this.sessions = sessions;
    }

    @Override
    public WebSocketHandler decorate(WebSocketHandler handler) {
        return new WebSocketHandlerDecorator(handler) {
            @Override
            public void afterConnectionEstablished(org.springframework.web.socket.WebSocketSession session) throws Exception {
                if (!LocalAuthSessionRegistry.isValid(session)) { session.close(org.springframework.web.socket.CloseStatus.POLICY_VIOLATION); return; }
                sessions.registerWebSocket(session);
                try { super.afterConnectionEstablished(session); }
                catch (Exception error) { sessions.unregisterWebSocket(session); throw error; }
            }
            @Override public void afterConnectionClosed(org.springframework.web.socket.WebSocketSession session, org.springframework.web.socket.CloseStatus status) throws Exception {
                try { super.afterConnectionClosed(session, status); }
                finally { sessions.unregisterWebSocket(session); }
            }
        };
    }
}
