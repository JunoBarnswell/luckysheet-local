package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.boot.context.properties.bind.ConstructorBinding;

import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.nio.file.Path;

@ConfigurationProperties(prefix = "luckysheet.auth")
public record AuthProperties(String mode, String issuer, String audience, String jwksUrl, String bootstrapFile) {
    @ConstructorBinding
    public AuthProperties {
    }

    public AuthProperties(String issuer, String audience, String jwksUrl) {
        this(null, issuer, audience, jwksUrl, null);
    }

    public List<String> audiences() {
        if (audience == null || audience.isBlank()) return List.of();
        return Arrays.stream(audience.split(",")).map(String::trim).filter(value -> !value.isBlank()).toList();
    }

    public AuthMode authMode() {
        String configured = mode == null || mode.isBlank() ? "local" : mode.trim().toLowerCase(Locale.ROOT);
        try {
            return AuthMode.valueOf(configured.toUpperCase(Locale.ROOT));
        } catch (IllegalArgumentException error) {
            throw new IllegalStateException("LUCKYSHEET_AUTH_MODE must be local or oidc");
        }
    }

    public Path bootstrapPath() {
        String configured = bootstrapFile == null || bootstrapFile.isBlank()
                ? "data/bootstrap-token"
                : bootstrapFile.trim();
        return Path.of(configured).toAbsolutePath().normalize();
    }

    public enum AuthMode {
        LOCAL,
        OIDC
    }
}
