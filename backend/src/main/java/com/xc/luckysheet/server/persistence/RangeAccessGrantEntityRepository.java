package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.Collection;
import java.util.List;

public interface RangeAccessGrantEntityRepository extends JpaRepository<RangeAccessGrantEntity, RangeAccessGrantEntity.Id> {
    @Query("select g from RangeAccessGrantEntity g where g.id.regionId in :regionIds")
    List<RangeAccessGrantEntity> findForRegions(@Param("regionIds") Collection<String> regionIds);

    @Modifying
    @Query("delete from RangeAccessGrantEntity g where g.id.regionId = :regionId")
    void deleteForRegion(@Param("regionId") String regionId);
}
