package com.xc.luckysheet.server.config;

import com.xc.luckysheet.server.ws.OperationWebSocketHandler;
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

    public WebSocketConfig(OperationWebSocketHandler handler, WebSocketProperties properties) {
        this.handler = handler;
        this.properties = properties;
    }

    @Override
    public void registerWebSocketHandlers(WebSocketHandlerRegistry registry) {
        registry.addHandler(handler, "/ws")
                .setAllowedOrigins(properties.origins().toArray(String[]::new));
    }
}
