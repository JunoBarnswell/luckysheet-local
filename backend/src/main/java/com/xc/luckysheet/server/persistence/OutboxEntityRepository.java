package com.xc.luckysheet.server.persistence;

import org.springframework.data.domain.Pageable;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Lock;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import jakarta.persistence.LockModeType;
import java.time.Instant;
import java.util.List;
import java.util.Optional;
import java.util.UUID;

public interface OutboxEntityRepository extends JpaRepository<OutboxEntity, UUID> {
    Optional<OutboxEntity> findByUnitIdAndRevision(String unitId, long revision);

    @Lock(LockModeType.PESSIMISTIC_WRITE)
    @Query("select o from OutboxEntity o where o.publishedAt is null and o.nextAttemptAt <= :now and (o.leaseUntil is null or o.leaseUntil < :now) order by o.createdAt asc")
    List<OutboxEntity> findPendingForUpdate(@Param("now") Instant now, Pageable pageable);

    List<OutboxEntity> findByPublishedAtIsNotNullAndPublishedAtBefore(Instant cutoff);

    void deleteByUnitId(String unitId);
}
