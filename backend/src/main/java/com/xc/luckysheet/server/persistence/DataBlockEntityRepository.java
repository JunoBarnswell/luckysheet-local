package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

public interface DataBlockEntityRepository extends JpaRepository<DataBlockEntity, DataBlockEntity.Id> {
    void deleteByIdUnitId(String unitId);

    long countByIdUnitId(String unitId);

    @Query("select coalesce(sum(b.byteLength), 0) from DataBlockEntity b where b.id.unitId = :unitId")
    long totalBytesByUnitId(@Param("unitId") String unitId);
}
