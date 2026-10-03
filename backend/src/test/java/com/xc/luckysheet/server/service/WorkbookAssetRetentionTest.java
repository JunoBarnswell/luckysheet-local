package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.persistence.*;
import org.junit.jupiter.api.Test;
import java.time.Instant;
import java.util.Optional;
import java.util.Set;
import java.util.stream.Stream;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class WorkbookAssetRetentionTest {
    @Test void releaseUsesCurrentAndHistoricalReferencesAndRejectsCorruptHistory() {
        AssetEntityRepository assets = mock(AssetEntityRepository.class);
        WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
        CheckpointEntityRepository checkpoints = mock(CheckpointEntityRepository.class);
        OperationEntityRepository operations = mock(OperationEntityRepository.class);
        when(workbooks.findById("book")).thenReturn(Optional.of(new WorkbookEntity("book", "Book", "{}", 0, 0, Instant.now(), Instant.now())));
        when(workbooks.findForUpdate("book")).thenAnswer(invocation -> workbooks.findById("book"));
        when(checkpoints.streamSnapshotJsonByUnitId("book")).thenAnswer(invocation -> Stream.of("{\"assetId\":\"asset-retained\"}"));
        when(operations.streamEnvelopeJsonByUnitId("book")).thenAnswer(invocation -> Stream.of("{\"mutations\":[{\"before\":{\"assetId\":\"asset-undo\"}}]}"));
        var service = new WorkbookAssetService(assets, workbooks, mock(AccessControlService.class), mock(WorkbookLifecycleService.class), checkpoints, operations, new ObjectMapper());
        assertEquals("ASSET_REFERENCED", assertThrows(ServiceException.class, () -> service.release("book", "asset-retained", "editor")).code());
        assertEquals("ASSET_REFERENCED", assertThrows(ServiceException.class, () -> service.release("book", "asset-undo", "editor")).code());
        verify(assets, never()).deleteById(any());
        service.release("book", "asset-unused", "editor");
        verify(assets).deleteById(new AssetEntity.Id("book", "asset-unused"));
        clearInvocations(assets);
        when(checkpoints.streamSnapshotJsonByUnitId("book")).thenAnswer(invocation -> Stream.of("{} {}"));
        assertEquals("STORAGE_CORRUPT", assertThrows(ServiceException.class, () -> service.release("book", "asset-unused", "editor")).code());
        verify(assets, never()).deleteById(any());
    }
    @Test void clientReconciliationCannotDeleteAuthoritativelyRetainedAssets() {
        AssetEntityRepository assets = mock(AssetEntityRepository.class);
        WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
        CheckpointEntityRepository checkpoints = mock(CheckpointEntityRepository.class);
        OperationEntityRepository operations = mock(OperationEntityRepository.class);
        when(workbooks.findById("book")).thenReturn(Optional.of(new WorkbookEntity("book", "Book", "{\"assetId\":\"asset-live\"}", 0, 0, Instant.now(), Instant.now())));
        when(workbooks.findForUpdate("book")).thenAnswer(invocation -> workbooks.findById("book"));
        when(checkpoints.streamSnapshotJsonByUnitId("book")).thenAnswer(invocation -> Stream.empty());
        when(operations.streamEnvelopeJsonByUnitId("book")).thenAnswer(invocation -> Stream.empty());
        var retained = mock(AssetEntity.class); var unused = mock(AssetEntity.class);
        when(retained.getId()).thenReturn(new AssetEntity.Id("book", "asset-live"));
        when(unused.getId()).thenReturn(new AssetEntity.Id("book", "asset-unused"));
        when(assets.findAllByIdUnitId("book")).thenReturn(java.util.List.of(retained, unused));
        var service = new WorkbookAssetService(assets, workbooks, mock(AccessControlService.class), mock(WorkbookLifecycleService.class), checkpoints, operations, new ObjectMapper());
        service.reconcile("book", Set.of(), "editor");
        verify(assets, never()).delete(retained);
        verify(assets).delete(unused);
    }
}
