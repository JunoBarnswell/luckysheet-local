package com.xc.luckysheet.server.coordination;

import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.store.OutboxRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;

import java.time.Duration;
import java.time.Instant;
import java.util.List;

/** Reliably drains committed ORM outbox rows into Redis. */
@Component
@ConditionalOnProperty(prefix = "luckysheet.coordination", name = "redis-enabled", havingValue = "true")
public class CoordinationOutboxPublisher {
    private static final Logger LOGGER = LoggerFactory.getLogger(CoordinationOutboxPublisher.class);
    private final WorkbookStore store;
    private final RedisCoordinationPublisher redis;
    private final CoordinationProperties properties;
    private volatile Instant lastCleanup = Instant.EPOCH;

    public CoordinationOutboxPublisher(WorkbookStore store, RedisCoordinationPublisher redis, CoordinationProperties properties) {
        this.store = store;
        this.redis = redis;
        this.properties = properties;
    }

    @Scheduled(fixedDelayString = "${luckysheet.coordination.outbox-poll-interval-ms:1000}")
    public void publishPending() {
        Instant now = Instant.now();
        List<OutboxRow> rows = store.claimOutbox(now, now.plus(properties.outboxLease()), properties.outboxBatchSize());
        for (OutboxRow row : rows) {
            try {
                redis.publishRevision(row);
                store.markOutboxPublished(row.eventId(), Instant.now());
            } catch (RuntimeException error) {
                Instant nextAttempt = Instant.now().plus(backoff(row.attempts()));
                store.releaseOutbox(row.eventId(), nextAttempt);
                LOGGER.warn("Revision event {} was not published; retrying at {}", row.operationId(), nextAttempt, error);
            }
        }
        if (now.isAfter(lastCleanup.plus(Duration.ofHours(1)))) {
            store.deletePublishedOutboxBefore(now.minus(Duration.ofDays(7)));
            lastCleanup = now;
        }
    }

    private Duration backoff(int attempts) {
        long seconds = 1L << Math.min(6, Math.max(0, attempts - 1));
        return Duration.ofSeconds(Math.min(60, seconds));
    }
}
