package com.xc.luckysheet.server.store;

import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class WorkbookDataBlockStoreTest {
    @Test
    void rejectsNewBlockWhenWorkbookByteQuotaWouldBeExceeded() {
        DataBlockEntityRepository blocks = mock(DataBlockEntityRepository.class);
        WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
        when(workbooks.findForUpdate("unit")).thenReturn(Optional.of(mock(WorkbookEntity.class)));
        when(blocks.findById(new com.xc.luckysheet.server.persistence.DataBlockEntity.Id("unit", "source", "block")))
                .thenReturn(Optional.empty());
        when(blocks.totalBytesByUnitId("unit")).thenReturn(90L);
        when(blocks.countByIdUnitId("unit")).thenReturn(1L);
        WorkbookDataBlockStore store = new WorkbookDataBlockStore(blocks, workbooks);
        Instant now = Instant.now();

        assertThrows(ServiceException.class, () -> store.upsertWithinQuota(
                new DataBlockRow("unit", "source", "block", "checksum", 11, new byte[11], now, now), 100, 10));

        verify(workbooks).findForUpdate("unit");
        verify(blocks, never()).save(org.mockito.ArgumentMatchers.any());
    }
}
