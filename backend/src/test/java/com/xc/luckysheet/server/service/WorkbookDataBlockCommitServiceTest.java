package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.DataBlockRow;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.junit.jupiter.api.Test;

import java.time.Instant;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.doThrow;
import static org.mockito.Mockito.inOrder;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.never;
import static org.mockito.Mockito.when;

class WorkbookDataBlockCommitServiceTest {
    @Test
    void rechecksEditorAccessAfterTheWorkbookLockAndBeforePersistence() {
        WorkbookDataBlockStore store = mock(WorkbookDataBlockStore.class);
        AccessControlService access = mock(AccessControlService.class);
        WorkbookLifecycleService lifecycle = mock(WorkbookLifecycleService.class);
        WorkbookDataBlockCommitService service = new WorkbookDataBlockCommitService(store, access, lifecycle);
        DataBlockRow row = new DataBlockRow("unit", "source", "block", "checksum", 1, new byte[] {1}, Instant.now(), Instant.now());
        doThrow(ServiceException.forbidden("Workbook access denied")).when(access).require("unit", "guest", WorkbookAclRole.EDITOR);

        assertThrows(ServiceException.class, () -> service.commit(row, 100, 10, "guest"));

        var order = inOrder(store, lifecycle, access);
        order.verify(store).lockWorkbook("unit");
        order.verify(lifecycle).requireActive("unit");
        order.verify(access).require("unit", "guest", WorkbookAclRole.EDITOR);
        org.mockito.Mockito.verify(store, never()).upsertWithinQuota(org.mockito.ArgumentMatchers.any(), org.mockito.ArgumentMatchers.anyLong(), org.mockito.ArgumentMatchers.anyLong());
    }
}
