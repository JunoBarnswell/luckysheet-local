package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

import java.time.Duration;

/** Configuration for the managed Rust workbook kernel process. */
@ConfigurationProperties(prefix = "luckysheet.kernel")
public record KernelHostProperties(
        String executable,
        String workingDirectory,
        String taskDirectory,
        Duration startupTimeout,
        Duration requestTimeout,
        int maxPayloadBytes,
        int maxPageBytes,
        int protocolVersion,
        int manifestVersion
) {
    public KernelHostProperties {
        executable = executable == null || executable.isBlank()
                ? defaultExecutable() : executable.trim();
        startupTimeout = startupTimeout == null ? Duration.ofSeconds(10) : startupTimeout;
        requestTimeout = requestTimeout == null ? Duration.ofMinutes(2) : requestTimeout;
        if (requestTimeout.isNegative() || requestTimeout.isZero()) throw new IllegalArgumentException("luckysheet.kernel.request-timeout must be positive");
        if (startupTimeout.isNegative() || startupTimeout.isZero()) {
            throw new IllegalArgumentException("luckysheet.kernel.startup-timeout must be positive");
        }
        if (maxPayloadBytes < 1024 || maxPayloadBytes > 16 * 1024 * 1024) {
            throw new IllegalArgumentException("luckysheet.kernel.max-payload-bytes must be between 1 KiB and 16 MiB");
        }
        if (maxPageBytes < 1024 || maxPageBytes > 1024 * 1024) {
            throw new IllegalArgumentException("luckysheet.kernel.max-page-bytes must be between 1 KiB and 1 MiB");
        }
        if (protocolVersion != 1 || manifestVersion != 11) {
            throw new IllegalArgumentException("Only kernel protocol 1 and manifest 11 are supported");
        }
    }

    public static KernelHostProperties defaults() {
        return new KernelHostProperties(defaultExecutable(), null, null, Duration.ofSeconds(10), Duration.ofMinutes(2),
                16 * 1024 * 1024, 1024 * 1024, 1, 11);
    }

    private static String defaultExecutable() {
        return System.getProperty("os.name", "").toLowerCase().contains("win")
                ? "target/release/workbook-kernel-host.exe" : "target/release/workbook-kernel-host";
    }
}
