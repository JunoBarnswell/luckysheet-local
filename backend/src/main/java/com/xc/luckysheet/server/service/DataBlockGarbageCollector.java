package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

/** Deletes expired staging bytes only after authoritative snapshot/history reference checks. */
@Service
public class DataBlockGarbageCollector {
    private final DataBlockEntityRepository blocks;
    private final WorkbookDataBlockStore store;
    private final WorkbookDataBlockReferenceGuard references;
    public DataBlockGarbageCollector(DataBlockEntityRepository blocks, WorkbookDataBlockStore store, WorkbookDataBlockReferenceGuard references) { this.blocks = blocks; this.store = store; this.references = references; }
    @Transactional public boolean collect(com.xc.luckysheet.server.contract.DataBlockMetadata block) {
        store.lockWorkbook(block.unitId());
        var current = store.findMetadata(block.unitId(), block.sourceId(), block.blockId());
        if (current.isEmpty() || !current.get().updatedAt().isBefore(java.time.Instant.now().minusSeconds(86400))) return false;
        references.requireUnreferenced(block.unitId(), block.sourceId(), block.blockId());
        store.delete(block.unitId(), block.sourceId(), block.blockId()); return true;
    }
}
