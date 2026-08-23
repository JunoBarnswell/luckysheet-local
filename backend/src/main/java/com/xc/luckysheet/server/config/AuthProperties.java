package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

import java.util.Arrays;
import java.util.List;

@ConfigurationProperties(prefix = "luckysheet.auth")
public record AuthProperties(String issuer, String audience, String jwksUrl) {
    public List<String> audiences() {
        if (audience == null || audience.isBlank()) return List.of();
        return Arrays.stream(audience.split(",")).map(String::trim).filter(value -> !value.isBlank()).toList();
    }
}
