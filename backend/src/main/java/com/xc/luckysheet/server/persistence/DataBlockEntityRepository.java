package com.xc.luckysheet.server.persistence;

import com.xc.luckysheet.server.contract.DataBlockMetadata;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.Optional;

public interface DataBlockEntityRepository extends JpaRepository<DataBlockEntity, DataBlockEntity.Id> {
    @Query("select new com.xc.luckysheet.server.contract.DataBlockMetadata(b.id.unitId, b.id.sourceId, b.id.blockId, b.checksum, b.byteLength, b.updatedAt) "
            + "from DataBlockEntity b where b.id.unitId = :unitId and b.id.sourceId = :sourceId and b.id.blockId = :blockId")
    Optional<DataBlockMetadata> findMetadata(
            @Param("unitId") String unitId, @Param("sourceId") String sourceId, @Param("blockId") String blockId);

    @Modifying(flushAutomatically = true)
    @Query(value = """
            insert into workbook_data_block
                (unit_id, source_id, block_id, checksum, byte_length, content, created_at, updated_at)
            select :targetUnitId, source_id, block_id, checksum, byte_length, content, created_at, updated_at
            from workbook_data_block
            where unit_id = :sourceUnitId and source_id = :sourceId and block_id = :blockId
            """, nativeQuery = true)
    int copyToWorkbook(
            @Param("sourceUnitId") String sourceUnitId,
            @Param("targetUnitId") String targetUnitId,
            @Param("sourceId") String sourceId,
            @Param("blockId") String blockId
    );

    void deleteByIdUnitId(String unitId);

    long countByIdUnitId(String unitId);

    @Query("select coalesce(sum(b.byteLength), 0) from DataBlockEntity b where b.id.unitId = :unitId")
    long totalBytesByUnitId(@Param("unitId") String unitId);
}
