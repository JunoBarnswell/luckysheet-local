package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

import java.util.Arrays;
import java.util.List;

@ConfigurationProperties(prefix = "luckysheet.websocket")
public record WebSocketProperties(String allowedOrigins) {
    public List<String> origins() {
        if (allowedOrigins == null || allowedOrigins.isBlank()) return List.of();
        return Arrays.stream(allowedOrigins.split(",")).map(String::trim).filter(value -> !value.isBlank()).toList();
    }
}
