package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

import java.time.Duration;

@ConfigurationProperties(prefix = "luckysheet.share")
public record ShareProperties(Duration defaultLifetime, Duration maxLifetime) {
    public ShareProperties {
        if (defaultLifetime == null || defaultLifetime.isZero() || defaultLifetime.isNegative()) throw new IllegalStateException("Share default lifetime must be positive");
        if (maxLifetime == null || maxLifetime.isZero() || maxLifetime.isNegative()) throw new IllegalStateException("Share max lifetime must be positive");
        if (defaultLifetime.compareTo(maxLifetime) > 0) throw new IllegalStateException("Share default lifetime exceeds max lifetime");
    }
}
