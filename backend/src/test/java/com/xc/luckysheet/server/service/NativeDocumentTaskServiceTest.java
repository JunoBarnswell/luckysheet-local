package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.config.KernelHostProperties;
import com.xc.luckysheet.server.contract.CreateNativeDocumentTaskRequest;
import com.xc.luckysheet.server.contract.WorkbookImportResponse;
import com.xc.luckysheet.server.persistence.NativeDocumentTaskEntity;
import com.xc.luckysheet.server.persistence.NativeDocumentTaskEntityRepository;
import jakarta.persistence.EntityManager;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.transaction.PlatformTransactionManager;
import org.springframework.transaction.support.SimpleTransactionStatus;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.time.Duration;
import java.util.HashMap;
import java.util.HexFormat;
import java.util.Map;
import java.util.Optional;
import java.util.function.BooleanSupplier;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

class NativeDocumentTaskServiceTest {
    @TempDir Path root;
    private NativeDocumentTaskService service;
    private WorkbookCatalogService catalog;
    private Map<String, NativeDocumentTaskEntity> rows;

    @BeforeEach void setup() {
        rows = new HashMap<>();
        NativeDocumentTaskEntityRepository repository = mock(NativeDocumentTaskEntityRepository.class);
        when(repository.save(any())).thenAnswer(call -> {
            NativeDocumentTaskEntity row = call.getArgument(0); rows.put(row.getTaskId(), row); return row;
        });
        when(repository.findForUpdate(anyString(), anyString())).thenAnswer(call -> owned(call.getArgument(0), call.getArgument(1)));
        when(repository.findByTaskIdAndActorSubject(anyString(), anyString())).thenAnswer(call -> owned(call.getArgument(0), call.getArgument(1)));
        catalog = mock(WorkbookCatalogService.class);
        PlatformTransactionManager manager = mock(PlatformTransactionManager.class);
        when(manager.getTransaction(any())).thenAnswer(call -> new SimpleTransactionStatus());
        service = new NativeDocumentTaskService(repository, catalog,
                new KernelHostProperties("kernel", null, root.toString(), Duration.ofSeconds(10), Duration.ofMinutes(2),
                        16 * 1024 * 1024, 1024 * 1024, 1, 11), new ObjectMapper(), manager, mock(EntityManager.class));
    }

    @Test void streamsSequentialChunksThenPublishesOnceAndPersistsResult() throws Exception {
        byte[] data = {1, 2, 3, 4};
        String id = create(data);
        AtomicReference<BooleanSupplier> finalGate = new AtomicReference<>();
        assertEquals(2, service.upload(id, 0, new ByteArrayInputStream(new byte[]{1, 2}), "alice").uploadedBytes());
        assertEquals(4, service.upload(id, 2, new ByteArrayInputStream(new byte[]{3, 4}), "alice").uploadedBytes());
        when(catalog.importNativePath(any(), any(), any(), any(), eq("alice"), any())).thenAnswer(call -> {
            assertEquals(data.length, Files.size(call.getArgument(0, Path.class)));
            try (InputStream input = Files.newInputStream(call.getArgument(0, Path.class))) {
                for (byte value : data) assertEquals(value, input.read());
                assertEquals(-1, input.read());
            }
            finalGate.set(call.getArgument(5));
            assertFalse(finalGate.get().getAsBoolean());
            return new WorkbookImportResponse("book", 0, "digest", null, new ObjectMapper().createObjectNode(), null);
        });
        assertEquals("completed", service.commit(id, "alice").state());
        assertFalse(finalGate.get().getAsBoolean(), "The beforeCommit callback must permit its own completed marker");
        assertEquals("book", service.commit(id, "alice").result().unitId());
        verify(catalog, times(1)).importNativePath(any(), any(), any(), any(), eq("alice"), any());
        assertFalse(Files.exists(root.resolve("uploads").resolve(id)));
    }

    @Test void rejectsOtherActorsIncompleteCommitAndWrongOffsets() throws Exception {
        String id = create(new byte[]{1, 2});
        assertCode("NOT_FOUND", () -> service.read(id, "mallory"));
        assertCode("NOT_FOUND", () -> service.upload(id, 0, new ByteArrayInputStream(new byte[]{1}), "mallory"));
        assertCode("NOT_FOUND", () -> service.cancel(id, "mallory"));
        assertCode("UPLOAD_INCOMPLETE", () -> service.commit(id, "alice"));
        assertCode("UPLOAD_OFFSET_MISMATCH", () -> service.upload(id, 1, new ByteArrayInputStream(new byte[]{1}), "alice"));
        verifyNoInteractions(catalog);
    }

    @Test void rejectsHashMismatchAndRestoresTheLastDurableOffset() throws Exception {
        String id = create(new byte[]{1, 2});
        service.upload(id, 0, new ByteArrayInputStream(new byte[]{1}), "alice");
        assertCode("UPLOAD_HASH_MISMATCH", () -> service.upload(id, 1, new ByteArrayInputStream(new byte[]{9}), "alice"));
        assertEquals(1, service.read(id, "alice").uploadedBytes());
        assertEquals(1, Files.size(root.resolve("uploads").resolve(id).resolve("source.xlsx")));
        assertEquals(2, service.upload(id, 1, new ByteArrayInputStream(new byte[]{2}), "alice").uploadedBytes());
    }

    @Test void rejectsOversizedChunkWithoutMaterializingIt() {
        String id = service.create(new CreateNativeDocumentTaskRequest("source.xlsx", null, null, null,
                NativeDocumentTaskService.MAX_CHUNK_BYTES + 1, "0".repeat(64)), "alice").taskId();
        InputStream input = new InputStream() {
            long remaining = NativeDocumentTaskService.MAX_CHUNK_BYTES + 1;
            @Override public int read() { return remaining-- > 0 ? 0 : -1; }
            @Override public int read(byte[] buffer, int offset, int length) {
                if (remaining <= 0) return -1;
                int count = (int) Math.min(remaining, length); remaining -= count; return count;
            }
        };
        assertCode("UPLOAD_LIMIT_EXCEEDED", () -> service.upload(id, 0, input, "alice"));
        assertEquals(0, service.read(id, "alice").uploadedBytes());
    }

    @Test void cancellationDuringNativeParsingPreventsPublicationAndCleansInput() throws Exception {
        String id = create(new byte[]{1});
        service.upload(id, 0, new ByteArrayInputStream(new byte[]{1}), "alice");
        when(catalog.importNativePath(any(), any(), any(), any(), eq("alice"), any())).thenAnswer(call -> {
            assertEquals("cancelled", service.cancel(id, "alice").state());
            assertTrue(Files.exists(call.getArgument(0, Path.class)));
            assertEquals("cancelled", service.cancel(id, "alice").state());
            assertTrue(call.getArgument(5, BooleanSupplier.class).getAsBoolean());
            throw new ServiceException("IMPORT_CANCELLED", 409, "cancelled");
        });
        assertCode("IMPORT_CANCELLED", () -> service.commit(id, "alice"));
        assertEquals("cancelled", service.read(id, "alice").state());
        assertFalse(Files.exists(root.resolve("uploads").resolve(id)));
        assertCode("IMPORT_TASK_STATE_CONFLICT", () -> service.commit(id, "alice"));
    }

    @Test void rejectsDocumentOverOneGiBAndClientPathsBeforeCreatingTask() {
        assertCode("UPLOAD_LIMIT_EXCEEDED", () -> service.create(new CreateNativeDocumentTaskRequest("source.xlsx",
                null, null, null, NativeDocumentTaskService.MAX_DOCUMENT_BYTES + 1, "0".repeat(64)), "alice"));
        assertCode("VALIDATION_ERROR", () -> service.create(new CreateNativeDocumentTaskRequest("../source.xlsx",
                null, null, null, 1, "0".repeat(64)), "alice"));
        assertTrue(rows.isEmpty());
    }

    @Test void nativeFailureIsDurableAndCannotBeRecommitted() throws Exception {
        String id = create(new byte[]{1});
        service.upload(id, 0, new ByteArrayInputStream(new byte[]{1}), "alice");
        when(catalog.importNativePath(any(), any(), any(), any(), eq("alice"), any()))
                .thenThrow(new ServiceException("UNSUPPORTED_FEATURE", 422, "unsupported workbook behavior"));
        assertCode("UNSUPPORTED_FEATURE", () -> service.commit(id, "alice"));
        assertEquals("failed", service.read(id, "alice").state());
        assertEquals("UNSUPPORTED_FEATURE", service.read(id, "alice").errorCode());
        assertCode("IMPORT_TASK_STATE_CONFLICT", () -> service.commit(id, "alice"));
        assertFalse(Files.exists(root.resolve("uploads").resolve(id)));
    }

    private String create(byte[] bytes) throws Exception {
        return service.create(new CreateNativeDocumentTaskRequest("source.xlsx", null, null, null, bytes.length,
                HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(bytes))), "alice").taskId();
    }
    private Optional<NativeDocumentTaskEntity> owned(String id, String actor) {
        return Optional.ofNullable(rows.get(id)).filter(row -> row.getActorSubject().equals(actor));
    }
    private static void assertCode(String code, org.junit.jupiter.api.function.Executable executable) {
        assertEquals(code, assertThrows(ServiceException.class, executable).code());
    }
}
