package com.xc.luckysheet.server.store;

import com.xc.luckysheet.server.persistence.DataBlockEntity;
import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.service.ServiceException;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;

import java.util.Optional;

@Repository
public class WorkbookDataBlockStore {
    private final DataBlockEntityRepository blocks;
    private final WorkbookEntityRepository workbooks;

    public WorkbookDataBlockStore(DataBlockEntityRepository blocks, WorkbookEntityRepository workbooks) {
        this.blocks = blocks;
        this.workbooks = workbooks;
    }

    public Optional<DataBlockRow> find(String unitId, String sourceId, String blockId) {
        return blocks.findById(new DataBlockEntity.Id(unitId, sourceId, blockId)).map(this::row);
    }

    @Transactional
    public DataBlockRow upsertWithinQuota(DataBlockRow row, long maximumBytes, long maximumBlocks) {
        workbooks.findForUpdate(row.unitId()).orElseThrow(() -> ServiceException.notFound("Workbook not found"));
        Optional<DataBlockEntity> existing = blocks.findById(new DataBlockEntity.Id(row.unitId(), row.sourceId(), row.blockId()));
        boolean isNew = existing.isEmpty();
        DataBlockEntity entity = existing.orElseGet(() -> new DataBlockEntity(row.unitId(), row.sourceId(), row.blockId(), row.checksum(),
                row.byteLength(), row.content().clone(), row.createdAt(), row.updatedAt()));
        long resultingBytes = blocks.totalBytesByUnitId(row.unitId()) - (isNew ? 0 : entity.getByteLength()) + row.byteLength();
        long resultingBlocks = blocks.countByIdUnitId(row.unitId()) + (isNew ? 1 : 0);
        if (resultingBytes > maximumBytes || resultingBlocks > maximumBlocks) {
            throw ServiceException.validation("Workbook data block quota exceeded");
        }
        if (!isNew && entity.getChecksum().equals(row.checksum()) && entity.getByteLength() == row.byteLength()) {
            return this.row(entity);
        }
        entity.update(row.checksum(), row.byteLength(), row.content().clone(), row.updatedAt());
        blocks.save(entity);
        return this.row(entity);
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
