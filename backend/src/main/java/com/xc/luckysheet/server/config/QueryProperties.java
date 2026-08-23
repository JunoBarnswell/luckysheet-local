package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

import java.time.Duration;
import java.util.Map;

@ConfigurationProperties(prefix = "luckysheet.query")
public record QueryProperties(
        boolean enabled,
        int maxRows,
        int maxColumns,
        int maxResponseBytes,
        Duration timeout,
        int workerThreads,
        Map<String, QuerySource> sources
) {
    public QueryProperties {
        if (maxRows < 1 || maxRows > 1_000_000) throw new IllegalStateException("Query max rows must be between 1 and 1000000");
        if (maxColumns < 1 || maxColumns > 10_000) throw new IllegalStateException("Query max columns must be between 1 and 10000");
        if (maxResponseBytes < 1_024 || maxResponseBytes > 100_000_000) throw new IllegalStateException("Query response size is invalid");
        if (timeout == null || timeout.isZero() || timeout.isNegative() || timeout.compareTo(Duration.ofMinutes(5)) > 0) {
            throw new IllegalStateException("Query timeout must be between 1ms and 5m");
        }
        if (workerThreads < 1 || workerThreads > 64) throw new IllegalStateException("Query worker thread count is invalid");
        sources = sources == null ? Map.of() : Map.copyOf(sources);
    }

    public QuerySource requireSource(String sourceRef, String connectorId) {
        QuerySource source = sources.get(sourceRef);
        if (source == null) throw new IllegalArgumentException("Unknown server query sourceRef");
        if (!source.kind().equalsIgnoreCase(connectorId)) {
            throw new IllegalArgumentException("Query connectorId does not match the configured source");
        }
        return source;
    }
}
