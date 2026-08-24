package com.xc.luckysheet.server.store;

import com.xc.luckysheet.server.contract.DataBlockMetadata;
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

    /** Acquires the canonical workbook write lock for authorization and block persistence. */
    public void lockWorkbook(String unitId) {
        workbooks.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found"));
    }

    /**
     * Persists under a caller-owned workbook write lock. Write callers receive
     * metadata only; binary content is exposed exclusively by {@link #find}.
     */
    public DataBlockMetadata upsertWithinQuota(DataBlockRow row, long maximumBytes, long maximumBlocks) {
        Optional<DataBlockEntity> existing = blocks.findById(new DataBlockEntity.Id(row.unitId(), row.sourceId(), row.blockId()));
        boolean isNew = existing.isEmpty();
        if (!isNew && entityMatches(existing.get(), row)) {
            return metadata(existing.get());
        }
        DataBlockEntity entity = existing.orElseGet(() -> new DataBlockEntity(row.unitId(), row.sourceId(), row.blockId(), row.checksum(),
                row.byteLength(), row.content().clone(), row.createdAt(), row.updatedAt()));
        long resultingBytes = blocks.totalBytesByUnitId(row.unitId()) - (isNew ? 0 : entity.getByteLength()) + row.byteLength();
        long resultingBlocks = blocks.countByIdUnitId(row.unitId()) + (isNew ? 1 : 0);
        if (resultingBytes > maximumBytes || resultingBlocks > maximumBlocks) {
            throw ServiceException.validation("Workbook data block quota exceeded");
        }
        entity.update(row.checksum(), row.byteLength(), row.content().clone(), row.updatedAt());
        blocks.save(entity);
        return metadata(entity);
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

    private static boolean entityMatches(DataBlockEntity entity, DataBlockRow row) {
        return entity.getChecksum().equals(row.checksum()) && entity.getByteLength() == row.byteLength();
    }

    private static DataBlockMetadata metadata(DataBlockEntity entity) {
        return new DataBlockMetadata(entity.getId().getUnitId(), entity.getId().getSourceId(), entity.getId().getBlockId(), entity.getChecksum(),
                entity.getByteLength(), entity.getUpdatedAt());
    }
}
