package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.time.Instant;
import java.util.List;
import java.util.Optional;
import java.util.UUID;

public interface ShareEntityRepository extends JpaRepository<ShareEntity, UUID> {
    List<ShareEntity> findByUnitIdOrderByCreatedAtDesc(String unitId);

    Optional<ShareEntity> findByUnitIdAndShareId(String unitId, UUID shareId);

    @Query("select s from ShareEntity s where s.unitId = :unitId and s.shareId = :shareId and s.revokedAt is null and s.expiresAt > :now")
    Optional<ShareEntity> findActiveForWorkbook(@Param("unitId") String unitId, @Param("shareId") UUID shareId, @Param("now") Instant now);

    @Query("select s from ShareEntity s where s.shareId = :shareId and s.revokedAt is null and s.expiresAt > :now")
    Optional<ShareEntity> findActive(@Param("shareId") UUID shareId, @Param("now") Instant now);

    void deleteByUnitId(String unitId);
}
