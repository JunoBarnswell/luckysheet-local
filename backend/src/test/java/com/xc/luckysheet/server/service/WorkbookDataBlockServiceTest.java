package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.verifyNoInteractions;

class WorkbookDataBlockServiceTest {
    private final WorkbookDataBlockStore store = mock(WorkbookDataBlockStore.class);
    private final WorkbookDataBlockCommitService commits = mock(WorkbookDataBlockCommitService.class);
    private final WorkbookDataBlockService service = new WorkbookDataBlockService(
            store, commits, mock(AccessControlService.class), mock(WorkbookLifecycleService.class));

    @Test
    void rejectsDeclaredOversizedBodyBeforeOpeningStream() {
        ServiceException error = assertThrows(ServiceException.class, () -> service.put(
                "unit", "source", "block", "checksum", WorkbookDataBlockService.MAX_BLOCK_BYTES + 1L,
                () -> { throw new AssertionError("body must not be opened"); }, "editor"));

        assertEquals(400, error.status());
        verifyNoInteractions(store);
        verifyNoInteractions(commits);
    }

    @Test
    void boundsChunkedBodyWhileReading() {
        byte[] oversized = new byte[WorkbookDataBlockService.MAX_BLOCK_BYTES + 1];

        ServiceException error = assertThrows(ServiceException.class, () -> service.put(
                "unit", "source", "block", "checksum", -1,
                () -> new ByteArrayInputStream(oversized), "editor"));

        assertEquals(400, error.status());
        verifyNoInteractions(store);
        verifyNoInteractions(commits);
    }

    @Test
    void delegatesOnlyVerifiedBytesToTheLockedCommitBoundary() {
        byte[] content = "content".getBytes(java.nio.charset.StandardCharsets.UTF_8);
        String checksum = "ed7002b439e9ac845f22357d822bac1444730f18b3f9d4e1f2b05b2d7d3a7f2a";
        var expected = new com.xc.luckysheet.server.contract.DataBlockMetadata("unit", "source", "block", checksum, content.length, java.time.Instant.now());
        when(commits.commit(org.mockito.ArgumentMatchers.any(), org.mockito.ArgumentMatchers.anyLong(), org.mockito.ArgumentMatchers.anyLong(), org.mockito.ArgumentMatchers.eq("editor"))).thenReturn(expected);

        assertEquals(expected, service.put("unit", "source", "block", checksum, content.length,
                () -> new ByteArrayInputStream(content), "editor"));
        verify(commits).commit(org.mockito.ArgumentMatchers.argThat(row -> row.content() == content),
                org.mockito.ArgumentMatchers.eq(WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_BYTES),
                org.mockito.ArgumentMatchers.eq(WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_COUNT), org.mockito.ArgumentMatchers.eq("editor"));
    }
}
