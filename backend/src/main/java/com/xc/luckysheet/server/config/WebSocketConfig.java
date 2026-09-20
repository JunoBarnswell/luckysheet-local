package com.xc.luckysheet.server.config;

import com.xc.luckysheet.server.ws.OperationWebSocketHandler;
import com.xc.luckysheet.server.security.LocalAuthSessionRegistry;
import com.xc.luckysheet.server.security.LocalAuthWebSocketHandlerDecoratorFactory;
import com.xc.luckysheet.server.security.WebSocketAuthenticationHandshakeHandler;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Configuration;
import org.springframework.web.socket.config.annotation.EnableWebSocket;
import org.springframework.web.socket.config.annotation.WebSocketConfigurer;
import org.springframework.web.socket.config.annotation.WebSocketHandlerRegistry;

@Configuration
@EnableWebSocket
@EnableConfigurationProperties(WebSocketProperties.class)
public class WebSocketConfig implements WebSocketConfigurer {
    private final OperationWebSocketHandler handler;
    private final WebSocketProperties properties;
    private final WebSocketAuthenticationHandshakeHandler authenticationHandler;
    private final LocalAuthSessionRegistry authSessions;

    public WebSocketConfig(
            OperationWebSocketHandler handler,
            WebSocketProperties properties,
            WebSocketAuthenticationHandshakeHandler authenticationHandler,
            LocalAuthSessionRegistry authSessions
    ) {
        this.handler = handler;
        this.properties = properties;
        this.authenticationHandler = authenticationHandler;
        this.authSessions = authSessions;
    }

    @Override
    public void registerWebSocketHandlers(WebSocketHandlerRegistry registry) {
        registry.addHandler(new LocalAuthWebSocketHandlerDecoratorFactory(authSessions).decorate(handler), "/ws")
                .setHandshakeHandler(authenticationHandler)
                .setAllowedOrigins(properties.origins().toArray(String[]::new));
    }
}
