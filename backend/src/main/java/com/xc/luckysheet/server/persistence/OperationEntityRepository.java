package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;
import org.springframework.data.domain.Pageable;

public interface OperationEntityRepository extends JpaRepository<OperationEntity, String> {
    Optional<OperationEntity> findByUnitIdAndActorSubjectAndClientSequence(String unitId, String actorSubject, long clientSequence);

    List<OperationEntity> findByUnitIdOrderByRevisionDesc(String unitId);

    List<OperationEntity> findByUnitIdAndRevisionLessThanOrderByRevisionDesc(String unitId, long revision, Pageable pageable);

    void deleteByUnitId(String unitId);
}
