package com.xc.luckysheet.server.config;

import org.junit.jupiter.api.Test;

import java.time.Duration;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;

class CoordinationPropertiesTest {
    private static CoordinationProperties defaults(boolean multiInstance, boolean redisEnabled, String redisUrl) {
        return new CoordinationProperties(
                multiInstance,
                redisEnabled,
                redisUrl,
                "luckysheet:coordination",
                Duration.ofSeconds(1),
                Duration.ofSeconds(30),
                100,
                Duration.ofSeconds(45)
        );
    }

    @Test
    void singleInstanceCanRunWithoutRedis() {
        assertDoesNotThrow(() -> defaults(false, false, ""));
    }

    @Test
    void multiInstanceRequiresEnabledRedis() {
        assertThrows(IllegalStateException.class, () -> defaults(true, false, ""));
        assertThrows(IllegalStateException.class, () -> defaults(true, true, ""));
        assertDoesNotThrow(() -> defaults(true, true, "redis://127.0.0.1:6379"));
    }
}
