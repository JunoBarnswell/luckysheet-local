package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.DataBlockMetadata;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.DataBlockRow;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

/**
 * The sole write boundary for uploaded data blocks. Streaming stays outside
 * this transaction; authorization, lifecycle, quota, and persistence share
 * the same workbook lock after the bytes have been fully verified.
 */
@Service
public class WorkbookDataBlockCommitService {
    private final WorkbookDataBlockStore store;
    private final AccessControlService access;
    private final WorkbookLifecycleService lifecycle;

    public WorkbookDataBlockCommitService(WorkbookDataBlockStore store, AccessControlService access,
                                          WorkbookLifecycleService lifecycle) {
        this.store = store;
        this.access = access;
        this.lifecycle = lifecycle;
    }

    @Transactional
    public DataBlockMetadata commit(DataBlockRow row, long maximumBytes, long maximumBlocks, String actor) {
        store.lockWorkbook(row.unitId());
        lifecycle.requireActive(row.unitId());
        access.require(row.unitId(), actor, WorkbookAclRole.EDITOR);
        return store.upsertWithinQuota(row, maximumBytes, maximumBlocks);
    }

    @Transactional
    public void delete(String unitId, String sourceId, String blockId, String actor) {
        store.lockWorkbook(unitId);
        lifecycle.requireActive(unitId);
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        store.delete(unitId, sourceId, blockId);
    }
}
