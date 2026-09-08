package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.config.KernelHostProperties;
import jakarta.annotation.PreDestroy;
import org.springframework.stereotype.Component;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

/**
 * Single canonical Java transport to the managed Rust workbook kernel.
 * Calls are serialized because the child process owns one ordered workbook
 * transaction stream. No Java semantic fallback is permitted.
 */
@Component
public final class KernelHostClient {
    private static final int HEADER_BYTES = 4;
    private final ObjectMapper mapper;
    private final KernelHostProperties properties;
    private Process process;
    private BufferedInputStream input;
    private BufferedOutputStream output;
    private boolean initialized;
    private final java.util.concurrent.ExecutorService readers = java.util.concurrent.Executors.newVirtualThreadPerTaskExecutor();

    public KernelHostClient(ObjectMapper mapper, KernelHostProperties properties) {
        this.mapper = mapper;
        this.properties = properties;
    }

    public synchronized JsonNode call(String operation, JsonNode params) {
        if (operation == null || operation.isBlank()) {
            throw new KernelHostException("INVALID_REQUEST", "Kernel operation is required", "operation", "correct-input");
        }
        ensureStarted();
        ObjectNode request = mapper.createObjectNode();
        request.put("protocolVersion", properties.protocolVersion());
        request.put("requestId", UUID.randomUUID().toString());
        request.put("operation", operation);
        request.set("params", params == null ? mapper.createObjectNode() : params);
        byte[] encoded;
        try {
            encoded = mapper.writeValueAsBytes(request);
        } catch (IOException error) {
            throw unavailable("Unable to encode kernel request", error);
        }
        if (encoded.length > properties.maxPayloadBytes()) {
            throw new KernelHostException("PAYLOAD_TOO_LARGE", "Kernel request exceeds payload limit", operation, "reduce-request-size");
        }
        try {
            writeFrame(encoded);
            JsonNode response = mapper.readTree(readFrameWithDeadline(properties.requestTimeout()));
            validateResponse(request.path("requestId").asText(), response);
            if (!response.path("ok").asBoolean(false)) throw fromError(response.path("error"));
            return response.path("result").isMissingNode() ? mapper.createObjectNode() : response.path("result");
        } catch (KernelHostException error) {
            if ("PROTOCOL_ERROR".equals(error.code()) || "TIMEOUT".equals(error.code())) stopProcess();
            throw error;
        } catch (Exception error) {
            stopProcess();
            throw unavailable("Kernel host transport failed", error);
        }
    }

    public int maxPageBytes() { return properties.maxPageBytes(); }

    /** Discards staged native state after a failed database transaction. */
    public synchronized void abortTransaction() { stopProcess(); }

    /** Releases one workbook context without starting a stopped host solely to close it. */
    public synchronized void closeWorkbookContext(String unitId) {
        if (unitId == null || unitId.isBlank()) throw new IllegalArgumentException("unitId is required");
        if (process == null || !process.isAlive()) return;
        try {
            call("close", mapper.createObjectNode().put("unitId", unitId));
        } catch (RuntimeException error) {
            stopProcess();
            throw error;
        }
    }

    private void ensureStarted() {
        if (process != null && process.isAlive()) {
            if (!initialized) initialize();
            return;
        }
        Path executable = Path.of(properties.executable()).toAbsolutePath().normalize();
        if (!Files.isRegularFile(executable) || !Files.isExecutable(executable)) {
            throw new KernelHostException("KERNEL_UNAVAILABLE", "Kernel host executable is missing or not executable",
                    executable.toString(), "build-and-install-workbook-kernel-host");
        }
        try {
            ProcessBuilder builder = new ProcessBuilder(executable.toString());
            if (properties.workingDirectory() != null && !properties.workingDirectory().isBlank()) {
                builder.directory(Path.of(properties.workingDirectory()).toAbsolutePath().normalize().toFile());
            }
            if (properties.taskDirectory() != null && !properties.taskDirectory().isBlank()) {
                Path directory = Path.of(properties.taskDirectory()).toAbsolutePath().normalize();
                Files.createDirectories(directory);
                builder.environment().put("KERNEL_TASK_DIRECTORY", directory.toRealPath().toString());
            }
            builder.redirectError(ProcessBuilder.Redirect.INHERIT);
            process = builder.start();
            input = new BufferedInputStream(process.getInputStream());
            output = new BufferedOutputStream(process.getOutputStream());
            initialized = false;
            initialize();
        } catch (IOException error) {
            stopProcess();
            throw unavailable("Unable to start kernel host", error);
        }
    }

    private void initialize() {
        JsonNode result = callWithoutInit("init", mapper.createObjectNode());
        if (result.path("manifestVersion").asInt(-1) != properties.manifestVersion()) {
            stopProcess();
            throw new KernelHostException("PROTOCOL_MISMATCH", "Kernel manifest version is not supported", "manifestVersion", "deploy-matching-kernel-host");
        }
        initialized = true;
    }

    private JsonNode callWithoutInit(String operation, JsonNode params) {
        ObjectNode request = mapper.createObjectNode();
        request.put("protocolVersion", properties.protocolVersion());
        request.put("requestId", UUID.randomUUID().toString());
        request.put("operation", operation);
        request.set("params", params);
        try {
            byte[] encoded = mapper.writeValueAsBytes(request);
            writeFrame(encoded);
            JsonNode response = mapper.readTree(readFrameWithDeadline(properties.startupTimeout()));
            validateResponse(request.path("requestId").asText(), response);
            if (!response.path("ok").asBoolean(false)) throw fromError(response.path("error"));
            return response.path("result");
        } catch (KernelHostException error) { throw error; }
        catch (Exception error) { stopProcess(); throw unavailable("Kernel initialization failed", error); }
    }

    private void validateResponse(String requestId, JsonNode response) {
        if (response == null || !response.isObject() || response.path("protocolVersion").asInt(-1) != properties.protocolVersion()
                || !requestId.equals(response.path("requestId").asText())) {
            throw new KernelHostException("PROTOCOL_ERROR", "Kernel response envelope is invalid", "response", "restart-matching-kernel-host");
        }
    }

    private KernelHostException fromError(JsonNode error) {
        return new KernelHostException(error.path("code").asText("KERNEL_ERROR"), error.path("message").asText("Kernel request failed"),
                error.path("object").isMissingNode() ? null : error.path("object").asText(null), error.path("recovery").asText("inspect-kernel-error"));
    }

    private void writeFrame(byte[] payload) throws IOException {
        output.write(ByteBuffer.allocate(HEADER_BYTES).order(ByteOrder.BIG_ENDIAN).putInt(payload.length).array());
        output.write(payload);
        output.flush();
    }

    private byte[] readFrameWithDeadline(java.time.Duration timeout) throws Exception {
        java.util.concurrent.Future<byte[]> read = readers.submit(this::readFrame);
        try { return read.get(timeout.toMillis(), TimeUnit.MILLISECONDS); }
        catch (java.util.concurrent.TimeoutException error) {
            if (process != null) process.destroyForcibly();
            read.cancel(true);
            throw new KernelHostException("TIMEOUT", "Kernel task exceeded its deadline", "workbook-kernel-host", "reduce-task-or-adjust-request-timeout");
        } catch (InterruptedException error) {
            if (process != null) process.destroyForcibly();
            read.cancel(true); Thread.currentThread().interrupt(); throw error;
        }
    }

    private byte[] readFrame() throws IOException {
        byte[] header = input.readNBytes(HEADER_BYTES);
        if (header.length != HEADER_BYTES) throw new IOException("Kernel closed before frame header");
        int length = ByteBuffer.wrap(header).order(ByteOrder.BIG_ENDIAN).getInt();
        if (length < 0 || length > properties.maxPayloadBytes()) throw new IOException("Kernel frame exceeds payload limit");
        byte[] payload = input.readNBytes(length);
        if (payload.length != length) throw new IOException("Kernel closed before frame payload");
        return payload;
    }

    private KernelHostException unavailable(String message, Exception cause) {
        return new KernelHostException("KERNEL_UNAVAILABLE", message + ": " + cause.getMessage(), "workbook-kernel-host", "restart-or-redeploy-kernel-host");
    }

    @PreDestroy
    public synchronized void close() { stopProcess(); }

    private void stopProcess() {
        initialized = false;
        if (process != null && process.isAlive()) process.destroyForcibly();
        if (output != null) try { output.close(); } catch (IOException ignored) { }
        if (input != null) try { input.close(); } catch (IOException ignored) { }
        if (process != null) {
            process.destroy();
            try { process.waitFor(1, TimeUnit.SECONDS); } catch (InterruptedException interrupted) { Thread.currentThread().interrupt(); }
            if (process.isAlive()) process.destroyForcibly();
        }
        output = null;
        input = null;
        process = null;
    }
}
