package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.domain.Pageable;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.Optional;

public interface CheckpointEntityRepository extends JpaRepository<CheckpointEntity, CheckpointEntity.Id> {
    @Query("select c from CheckpointEntity c where c.id.unitId = :unitId and c.id.revision = :revision")
    Optional<CheckpointEntity> findAtRevision(@Param("unitId") String unitId, @Param("revision") long revision);

    @Query("select c from CheckpointEntity c where c.id.unitId = :unitId and c.id.revision <= :revision order by c.id.revision desc")
    java.util.List<CheckpointEntity> findLatestAtOrBefore(@Param("unitId") String unitId, @Param("revision") long revision, Pageable pageable);
}
