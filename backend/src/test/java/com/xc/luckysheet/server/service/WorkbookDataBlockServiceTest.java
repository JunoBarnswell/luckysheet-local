package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verifyNoInteractions;

class WorkbookDataBlockServiceTest {
    private final WorkbookDataBlockStore store = mock(WorkbookDataBlockStore.class);
    private final WorkbookDataBlockService service = new WorkbookDataBlockService(
            store, mock(AccessControlService.class), mock(WorkbookLifecycleService.class));

    @Test
    void rejectsDeclaredOversizedBodyBeforeOpeningStream() {
        ServiceException error = assertThrows(ServiceException.class, () -> service.put(
                "unit", "source", "block", "checksum", WorkbookDataBlockService.MAX_BLOCK_BYTES + 1L,
                () -> { throw new AssertionError("body must not be opened"); }, "editor"));

        assertEquals(400, error.status());
        verifyNoInteractions(store);
    }

    @Test
    void boundsChunkedBodyWhileReading() {
        byte[] oversized = new byte[WorkbookDataBlockService.MAX_BLOCK_BYTES + 1];

        ServiceException error = assertThrows(ServiceException.class, () -> service.put(
                "unit", "source", "block", "checksum", -1,
                () -> new ByteArrayInputStream(oversized), "editor"));

        assertEquals(400, error.status());
        verifyNoInteractions(store);
    }
}
