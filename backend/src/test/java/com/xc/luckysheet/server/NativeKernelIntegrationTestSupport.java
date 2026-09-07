package com.xc.luckysheet.server;

import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.DynamicPropertyRegistry;
import org.springframework.test.context.DynamicPropertySource;
import org.springframework.test.context.TestPropertySource;

import java.nio.file.Files;
import java.nio.file.Path;

/** Every integration test starts the real native host; a missing build is a failure, never a skip. */
@SpringBootTest
@TestPropertySource(properties = {
        "DATABASE_USERNAME=sa", "DATABASE_PASSWORD=", "JPA_DDL_AUTO=validate",
        "FLYWAY_BASELINE_ON_MIGRATE=false", "AUTH_ISSUER=https://issuer.test", "AUTH_AUDIENCE=test",
        "AUTH_JWKS_URL=https://issuer.test/.well-known/jwks.json",
        "COORDINATION_MULTI_INSTANCE=false", "COORDINATION_REDIS_ENABLED=false"
})
public abstract class NativeKernelIntegrationTestSupport {
    @DynamicPropertySource
    static void nativeKernel(DynamicPropertyRegistry properties) {
        Path root = Path.of("").toAbsolutePath().normalize();
        while (root != null && !Files.isRegularFile(root.resolve("kernel/host/Cargo.toml"))) root = root.getParent();
        if (root == null) throw new IllegalStateException("Cannot locate repository kernel/host/Cargo.toml");
        String configured = System.getProperty("kernel.host.executable", System.getenv("KERNEL_HOST_EXECUTABLE"));
        String fileName = System.getProperty("os.name").toLowerCase().contains("win")
                ? "workbook-kernel-host.exe" : "workbook-kernel-host";
        Path executable = configured == null || configured.isBlank()
                ? root.resolve("target/debug").resolve(fileName) : Path.of(configured).toAbsolutePath().normalize();
        if (!Files.isRegularFile(executable)) throw new IllegalStateException(
                "Build the real native host before integration tests: cargo build -p workbook-kernel-host; missing " + executable);
        Path workingDirectory = root;
        properties.add("luckysheet.kernel.executable", executable::toString);
        properties.add("luckysheet.kernel.working-directory", workingDirectory::toString);
        properties.add("luckysheet.kernel.task-directory", () -> workingDirectory.resolve("target/native-integration-tasks").toString());
    }
}
