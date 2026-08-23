package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

import java.time.Duration;

@ConfigurationProperties(prefix = "luckysheet.coordination")
public record CoordinationProperties(
        boolean multiInstance,
        boolean redisEnabled,
        String redisUrl,
        String channel,
        Duration outboxPollInterval,
        Duration outboxLease,
        int outboxBatchSize,
        Duration presenceTtl
) {
    public CoordinationProperties {
        if (multiInstance && !redisEnabled) {
            throw new IllegalStateException("COORDINATION_REDIS_ENABLED must be true when COORDINATION_MULTI_INSTANCE is true");
        }
        if (redisEnabled && (redisUrl == null || redisUrl.isBlank())) {
            throw new IllegalStateException("COORDINATION_REDIS_URL must be configured when Redis coordination is enabled");
        }
        if (channel == null || channel.isBlank()) {
            throw new IllegalStateException("Redis coordination channel must be configured");
        }
        if (outboxPollInterval == null || outboxPollInterval.isZero() || outboxPollInterval.isNegative()) {
            throw new IllegalStateException("Outbox poll interval must be positive");
        }
        if (outboxLease == null || outboxLease.isZero() || outboxLease.isNegative()) {
            throw new IllegalStateException("Outbox lease must be positive");
        }
        if (outboxBatchSize < 1 || outboxBatchSize > 1000) {
            throw new IllegalStateException("Outbox batch size must be between 1 and 1000");
        }
        if (presenceTtl == null || presenceTtl.isZero() || presenceTtl.isNegative()) {
            throw new IllegalStateException("Presence TTL must be positive");
        }
    }
}
