package com.xc.luckysheet.server.store;

import com.xc.luckysheet.server.persistence.DataBlockEntity;
import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;

import java.util.Optional;

@Repository
public class WorkbookDataBlockStore {
    private final DataBlockEntityRepository blocks;

    public WorkbookDataBlockStore(DataBlockEntityRepository blocks) {
        this.blocks = blocks;
    }

    public Optional<DataBlockRow> find(String unitId, String sourceId, String blockId) {
        return blocks.findById(new DataBlockEntity.Id(unitId, sourceId, blockId)).map(this::row);
    }

    @Transactional
    public void upsert(DataBlockRow row) {
        DataBlockEntity entity = blocks.findById(new DataBlockEntity.Id(row.unitId(), row.sourceId(), row.blockId()))
                .orElseGet(() -> new DataBlockEntity(row.unitId(), row.sourceId(), row.blockId(), row.checksum(), row.byteLength(),
                        row.content().clone(), row.createdAt(), row.updatedAt()));
        entity.update(row.checksum(), row.byteLength(), row.content().clone(), row.updatedAt());
        blocks.save(entity);
    }

    @Transactional
    public int delete(String unitId, String sourceId, String blockId) {
        DataBlockEntity.Id id = new DataBlockEntity.Id(unitId, sourceId, blockId);
        if (!blocks.existsById(id)) return 0;
        blocks.deleteById(id);
        return 1;
    }

    private DataBlockRow row(DataBlockEntity entity) {
        return new DataBlockRow(entity.getId().getUnitId(), entity.getId().getSourceId(), entity.getId().getBlockId(), entity.getChecksum(),
                entity.getByteLength(), entity.getContent().clone(), entity.getCreatedAt(), entity.getUpdatedAt());
    }
}
