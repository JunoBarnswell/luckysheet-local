package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.config.KernelHostProperties;
import com.xc.luckysheet.server.contract.CreateNativeDocumentTaskRequest;
import com.xc.luckysheet.server.contract.NativeDocumentTaskResponse;
import com.xc.luckysheet.server.contract.WorkbookImportResponse;
import com.xc.luckysheet.server.persistence.NativeDocumentTaskEntity;
import com.xc.luckysheet.server.persistence.NativeDocumentTaskEntityRepository;
import jakarta.persistence.EntityManager;
import jakarta.persistence.LockModeType;
import org.springframework.stereotype.Service;
import org.springframework.transaction.PlatformTransactionManager;
import org.springframework.transaction.TransactionDefinition;
import org.springframework.transaction.support.TransactionTemplate;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.channels.FileChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.Comparator;
import java.util.HexFormat;
import java.util.Locale;
import java.util.UUID;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/** One file-backed native import lifecycle, shared by chunk and multipart entrypoints. */
@Service
public class NativeDocumentTaskService {
    public static final long MAX_DOCUMENT_BYTES = 1024L * 1024L * 1024L;
    public static final long MAX_CHUNK_BYTES = 8L * 1024L * 1024L;
    private static final int BUFFER_BYTES = 64 * 1024;

    private final NativeDocumentTaskEntityRepository tasks;
    private final WorkbookCatalogService catalog;
    private final KernelHostProperties properties;
    private final ObjectMapper mapper;
    private final TransactionTemplate transactions;
    private final EntityManager entityManager;
    /** File reader lifetimes only; all product state and cancellation authority remain in the database. */
    private final Set<String> activeReaders = ConcurrentHashMap.newKeySet();

    public NativeDocumentTaskService(NativeDocumentTaskEntityRepository tasks, WorkbookCatalogService catalog,
            KernelHostProperties properties, ObjectMapper mapper, PlatformTransactionManager transactionManager,
            EntityManager entityManager) {
        this.tasks = tasks; this.catalog = catalog; this.properties = properties; this.mapper = mapper;
        this.entityManager = entityManager;
        transactions = new TransactionTemplate(transactionManager);
        transactions.setPropagationBehavior(TransactionDefinition.PROPAGATION_REQUIRES_NEW);
        // The final cancellation lock must see DELETE committed during native parsing on MySQL as well.
        transactions.setIsolationLevel(TransactionDefinition.ISOLATION_READ_COMMITTED);
    }

    public NativeDocumentTaskResponse create(CreateNativeDocumentTaskRequest request, String actor) {
        if (request.sha256() == null || !request.sha256().matches("[0-9a-fA-F]{64}")) {
            throw ServiceException.validation("sha256 must contain exactly 64 hexadecimal characters");
        }
        return response(createTask(request.fileName(), request.name(), request.spaceId(), request.folderId(),
                request.byteLength(), request.sha256().toLowerCase(Locale.ROOT), actor));
    }

    public NativeDocumentTaskResponse read(String id, String actor) { return response(require(id, actor)); }

    public NativeDocumentTaskResponse upload(String id, long offset, InputStream input, String actor) {
        return upload(id, offset, input, actor, MAX_CHUNK_BYTES);
    }

    private NativeDocumentTaskResponse upload(String id, long offset, InputStream input, String actor, long limit) {
        return transactions.execute(status -> {
            NativeDocumentTaskEntity task = lock(id, actor);
            requireState(task, "uploading");
            if (offset != task.getUploadedBytes()) throw error("UPLOAD_OFFSET_MISMATCH", 409,
                    "Task " + id + " expects offset " + task.getUploadedBytes() + "; resend from that offset");
            Path path = inputPath(task);
            long written = 0;
            try {
                if (Files.size(path) != offset) throw error("UPLOAD_LENGTH_MISMATCH", 409,
                        "Task " + id + " file and durable offset differ; cancel and start a new task");
                try (OutputStream output = Files.newOutputStream(path, StandardOpenOption.APPEND)) {
                    byte[] buffer = new byte[BUFFER_BYTES];
                    int count;
                    while ((count = input.read(buffer)) != -1) {
                        if (count == 0) continue;
                        if (written + count > limit || offset + written + count > task.getByteLength()) {
                            throw error("UPLOAD_LIMIT_EXCEEDED", 413,
                                    "Task " + id + " chunk exceeds its byte limit; resend a bounded chunk");
                        }
                        output.write(buffer, 0, count);
                        written += count;
                    }
                }
                if (written == 0) throw error("UPLOAD_EMPTY_CHUNK", 400, "An upload chunk must contain bytes");
                try (FileChannel channel = FileChannel.open(path, StandardOpenOption.WRITE)) { channel.force(true); }
                if (offset + written == task.getByteLength()) verifyHash(task);
                task.uploaded(offset + written);
                tasks.save(task);
                return response(task);
            } catch (IOException | RuntimeException failure) {
                // File writes are not transactional: restore the durable offset before rejecting this request.
                try (FileChannel channel = FileChannel.open(path, StandardOpenOption.WRITE)) {
                    channel.truncate(offset); channel.force(true);
                } catch (IOException rollbackFailure) { failure.addSuppressed(rollbackFailure); }
                if (failure instanceof ServiceException service) throw service;
                throw error("UPLOAD_IO_FAILED", 500, "Task " + id + " could not persist its chunk; retry or cancel", failure);
            }
        });
    }

    public NativeDocumentTaskResponse commit(String id, String actor) {
        NativeDocumentTaskEntity ready = transactions.execute(status -> {
            NativeDocumentTaskEntity task = lock(id, actor);
            if ("completed".equals(task.getState())) return task;
            requireState(task, "uploading");
            if (task.getUploadedBytes() != task.getByteLength()) throw error("UPLOAD_INCOMPLETE", 409,
                    "Task " + id + " must upload all " + task.getByteLength() + " bytes before commit");
            verifyHash(task);
            task.importing();
            return tasks.save(task);
        });
        if ("completed".equals(ready.getState())) return response(ready);
        activeReaders.add(id);
        try {
            return transactions.execute(status -> {
                // Catalog participates in this transaction. Its final callback locks the task row only
                // after parsing, and publication plus the completed marker commit atomically.
                WorkbookImportResponse result = catalog.importNativePath(inputPath(ready), ready.getName(),
                        ready.getSpaceId(), ready.getFolderId(), actor,
                        () -> {
                            String state = lock(id, actor).getState();
                            return "cancelled".equals(state) || "failed".equals(state);
                        });
                NativeDocumentTaskEntity task = lock(id, actor);
                if (!"importing".equals(task.getState())) throw error("IMPORT_CANCELLED", 409,
                        "Task " + id + " was cancelled before publication");
                try { task.completed(mapper.writeValueAsString(result)); }
                catch (JsonProcessingException e) { throw error("IMPORT_RESULT_INVALID", 500,
                        "Task " + id + " result cannot be persisted; publication was aborted", e); }
                tasks.save(task);
                return response(task);
            });
        } catch (RuntimeException failure) {
            transactions.executeWithoutResult(status -> {
                NativeDocumentTaskEntity task = lock(id, actor);
                if ("importing".equals(task.getState())) {
                    task.failed(failure instanceof ServiceException service ? service.code() : "IMPORT_FAILED",
                            "Task " + id + " import failed: " + failure.getMessage() + "; create a new task to retry");
                    tasks.save(task);
                }
            });
            throw failure;
        } finally {
            try { cleanup(ready); }
            finally { activeReaders.remove(id); }
        }
    }

    public NativeDocumentTaskResponse cancel(String id, String actor) {
        boolean[] deferCleanup = {false};
        NativeDocumentTaskEntity task = transactions.execute(status -> {
            NativeDocumentTaskEntity current = lock(id, actor);
            // A completed publication cannot be converted into a cancellation after the fact.
            if ("completed".equals(current.getState()) || "failed".equals(current.getState())) return current;
            deferCleanup[0] = activeReaders.contains(id);
            current.cancelled();
            return tasks.save(current);
        });
        // Native readers can still hold this file on Windows. The committing request owns final cleanup.
        if (!deferCleanup[0]) cleanup(task);
        return response(task);
    }

    public WorkbookImportResponse importMultipart(MultipartFile file, String name, String spaceId,
            String folderId, String actor) {
        NativeDocumentTaskEntity task = createTask(file.getOriginalFilename(), name, spaceId, folderId,
                file.getSize(), null, actor);
        try (InputStream input = file.getInputStream()) {
            upload(task.getTaskId(), 0, input, actor, MAX_DOCUMENT_BYTES);
            return commit(task.getTaskId(), actor).result();
        } catch (IOException | RuntimeException failure) {
            cancel(task.getTaskId(), actor);
            if (failure instanceof ServiceException service) throw service;
            throw error("IMPORT_IO_FAILED", 500, "Multipart upload could not be read; retry with a new task", failure);
        }
    }

    private NativeDocumentTaskEntity createTask(String fileName, String name, String spaceId, String folderId,
            long length, String hash, String actor) {
        if (actor == null || actor.isBlank() || actor.length() > 500) throw ServiceException.forbidden("Registered actor is required");
        if (length < 1 || length > MAX_DOCUMENT_BYTES) throw error("UPLOAD_LIMIT_EXCEEDED", 413,
                "Native documents must contain between 1 byte and 1 GiB");
        if (fileName == null || fileName.isBlank() || fileName.length() > 500 || fileName.equals(".") || fileName.equals("..")
                || fileName.matches(".*[\\\\/:*?\"<>|\\p{Cntrl}].*") || fileName.endsWith(".") || fileName.endsWith(" ")) {
            throw ServiceException.validation("fileName must be a plain file name without path or control characters");
        }
        if (name != null && name.length() > 500 || spaceId != null && spaceId.length() > 200
                || folderId != null && folderId.length() > 200) throw ServiceException.validation("Import metadata exceeds its length limit");
        NativeDocumentTaskEntity task = new NativeDocumentTaskEntity(UUID.randomUUID().toString(), actor,
                fileName, name, spaceId, folderId, length, hash);
        try {
            Files.createDirectories(taskDirectory(task));
            Files.createFile(inputPath(task));
            return transactions.execute(status -> tasks.save(task));
        } catch (IOException | RuntimeException failure) {
            cleanup(task);
            if (failure instanceof ServiceException service) throw service;
            throw error("IMPORT_STORAGE_UNAVAILABLE", 503, "Cannot create native import task storage; check KERNEL_TASK_DIRECTORY", failure);
        }
    }

    private void verifyHash(NativeDocumentTaskEntity task) {
        try {
            Path file = inputPath(task);
            if (Files.size(file) != task.getByteLength()) throw error("UPLOAD_LENGTH_MISMATCH", 409,
                    "Task " + task.getTaskId() + " uploaded length differs from its declared length");
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            try (InputStream input = Files.newInputStream(file)) {
                byte[] buffer = new byte[BUFFER_BYTES]; int count;
                while ((count = input.read(buffer)) != -1) if (count > 0) digest.update(buffer, 0, count);
            }
            if (task.getSha256() != null && !HexFormat.of().formatHex(digest.digest()).equals(task.getSha256())) {
                throw error("UPLOAD_HASH_MISMATCH", 422,
                        "Task " + task.getTaskId() + " SHA-256 differs from the declared hash; resend correct bytes or cancel");
            }
        } catch (IOException e) { throw error("UPLOAD_IO_FAILED", 500, "Cannot verify task " + task.getTaskId(), e); }
        catch (NoSuchAlgorithmException e) { throw new IllegalStateException("SHA-256 is required by the JVM", e); }
    }

    private NativeDocumentTaskEntity require(String id, String actor) {
        return tasks.findByTaskIdAndActorSubject(id, actor).orElseThrow(() -> ServiceException.notFound("Import task was not found"));
    }
    private NativeDocumentTaskEntity lock(String id, String actor) {
        // Flush our own final completed marker before refresh. A request-scoped persistence context
        // may still cache the earlier importing row while another request has committed cancellation.
        entityManager.flush();
        NativeDocumentTaskEntity task = tasks.findForUpdate(id, actor)
                .orElseThrow(() -> ServiceException.notFound("Import task was not found"));
        entityManager.refresh(task, LockModeType.PESSIMISTIC_WRITE);
        return task;
    }
    private void requireState(NativeDocumentTaskEntity task, String expected) {
        if (!expected.equals(task.getState())) throw error("IMPORT_TASK_STATE_CONFLICT", 409,
                "Task " + task.getTaskId() + " is " + task.getState() + "; expected " + expected);
    }
    private Path taskDirectory(NativeDocumentTaskEntity task) {
        if (properties.taskDirectory() == null || properties.taskDirectory().isBlank()) throw error("IMPORT_STORAGE_UNAVAILABLE", 503,
                "KERNEL_TASK_DIRECTORY must be configured before creating import tasks");
        Path root = Path.of(properties.taskDirectory()).toAbsolutePath().normalize().resolve("uploads");
        Path directory = root.resolve(UUID.fromString(task.getTaskId()).toString()).normalize();
        if (!directory.startsWith(root)) throw error("IMPORT_PATH_INVALID", 500, "Task directory is outside import storage");
        return directory;
    }
    private Path inputPath(NativeDocumentTaskEntity task) { return taskDirectory(task).resolve(task.getFileName()); }
    private void cleanup(NativeDocumentTaskEntity task) {
        Path directory = taskDirectory(task);
        if (!Files.exists(directory)) return;
        try (var paths = Files.walk(directory)) {
            for (Path path : paths.sorted(Comparator.reverseOrder()).toList()) Files.deleteIfExists(path);
        } catch (IOException e) { throw error("IMPORT_CLEANUP_FAILED", 500,
                "Task " + task.getTaskId() + " temporary files could not be removed; retry DELETE after native processing stops", e); }
    }
    private NativeDocumentTaskResponse response(NativeDocumentTaskEntity task) {
        WorkbookImportResponse result = null;
        try {
            if (task.getResultJson() != null) result = mapper.readValue(task.getResultJson(), WorkbookImportResponse.class);
        } catch (JsonProcessingException e) { throw error("IMPORT_RESULT_INVALID", 500, "Stored import result is invalid", e); }
        return new NativeDocumentTaskResponse(task.getTaskId(), task.getState(), task.getUploadedBytes(),
                task.getByteLength(), result, task.getErrorCode(), task.getErrorMessage());
    }
    private static ServiceException error(String code, int status, String message) { return new ServiceException(code, status, message); }
    private static ServiceException error(String code, int status, String message, Throwable cause) {
        return new ServiceException(code, status, message, cause);
    }
}
