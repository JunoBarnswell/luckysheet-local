package com.xc.luckysheet.server.persistence;

import org.springframework.data.domain.PageRequest;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;

public interface AuditEntityRepository extends JpaRepository<AuditEntity, java.util.UUID> {
    default List<AuditEntity> findRecent(String unitId, int limit) {
        return findByUnitIdOrderByOccurredAtDesc(unitId, PageRequest.of(0, limit));
    }

    List<AuditEntity> findByUnitIdOrderByOccurredAtDesc(String unitId, org.springframework.data.domain.Pageable pageable);

    void deleteByUnitId(String unitId);
}
