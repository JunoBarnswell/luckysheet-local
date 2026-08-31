package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;
import org.springframework.data.domain.Pageable;

public interface OperationEntityRepository extends JpaRepository<OperationEntity, String> {
    Optional<OperationEntity> findByUnitIdAndActorSubjectAndClientSequence(String unitId, String actorSubject, long clientSequence);

    /**
     * The only replay query.  Both bounds are required so a checkpoint or a
     * revision reader can never turn a local replay into a full history scan.
     */
    List<OperationEntity> findByUnitIdAndRevisionGreaterThanAndRevisionLessThanEqualOrderByRevisionAsc(
            String unitId, long fromExclusive, long toInclusive, Pageable pageable);

    /** Bounded history page used by the human-facing revision log. */
    List<OperationEntity> findByUnitIdAndRevisionGreaterThanEqualAndRevisionLessThanOrderByRevisionDesc(
            String unitId, long fromInclusive, long toExclusive, Pageable pageable);

    void deleteByUnitId(String unitId);
}
