package com.xc.luckysheet.server.store;

import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertEquals;
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

        store.lockWorkbook("unit");
        assertThrows(ServiceException.class, () -> store.upsertWithinQuota(
                new DataBlockRow("unit", "source", "block", "checksum", 11, new byte[11], now, now), 100, 10));

        verify(workbooks).findForUpdate("unit");
        verify(blocks, never()).save(org.mockito.ArgumentMatchers.any());
    }

    @Test
    void returnsMetadataForAnIdempotentWriteWithoutReadingBlockContent() {
        DataBlockEntityRepository blocks = mock(DataBlockEntityRepository.class);
        WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
        com.xc.luckysheet.server.persistence.DataBlockEntity entity = mock(com.xc.luckysheet.server.persistence.DataBlockEntity.class);
        com.xc.luckysheet.server.persistence.DataBlockEntity.Id id = new com.xc.luckysheet.server.persistence.DataBlockEntity.Id("unit", "source", "block");
        Instant now = Instant.now();
        when(entity.getId()).thenReturn(id);
        when(entity.getChecksum()).thenReturn("checksum");
        when(entity.getByteLength()).thenReturn(4);
        when(entity.getUpdatedAt()).thenReturn(now);
        when(blocks.findById(id)).thenReturn(Optional.of(entity));
        WorkbookDataBlockStore store = new WorkbookDataBlockStore(blocks, workbooks);

        var metadata = store.upsertWithinQuota(new DataBlockRow("unit", "source", "block", "checksum", 4, new byte[] {1, 2, 3, 4}, now, now), 1, 1);

        assertEquals("block", metadata.blockId());
        verify(entity, never()).getContent();
        verify(blocks, never()).save(entity);
    }
}
