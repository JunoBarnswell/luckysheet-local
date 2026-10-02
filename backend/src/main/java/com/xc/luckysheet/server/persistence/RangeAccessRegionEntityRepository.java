package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.List;
import java.util.Optional;

public interface RangeAccessRegionEntityRepository extends JpaRepository<RangeAccessRegionEntity, String> {
    @Query("select r from RangeAccessRegionEntity r where r.unitId = :unitId order by r.sheetId, r.startRow, r.startColumn")
    List<RangeAccessRegionEntity> findAllForWorkbook(@Param("unitId") String unitId);

    @Query("select r from RangeAccessRegionEntity r where r.unitId = :unitId and r.id = :regionId")
    Optional<RangeAccessRegionEntity> findForWorkbook(@Param("unitId") String unitId, @Param("regionId") String regionId);
}
