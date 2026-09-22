package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.List;
import java.util.Optional;
import java.util.stream.Stream;
import org.springframework.data.domain.Pageable;

public interface OperationEntityRepository extends JpaRepository<OperationEntity, String> {
    Optional<OperationEntity> findByUnitIdAndActorSubjectAndClientSessionIdAndClientSequence(String unitId, String actorSubject, String clientSessionId, long clientSequence);

    List<OperationEntity> findByUnitIdOrderByRevisionDesc(String unitId);

    List<OperationEntity> findByUnitIdAndRevisionGreaterThanAndRevisionLessThanEqualOrderByRevisionAsc(String unitId, long afterRevision, long throughRevision);

    List<OperationEntity> findByUnitIdAndRevisionLessThanOrderByRevisionDesc(String unitId, long revision, Pageable pageable);

    void deleteByUnitId(String unitId);

    @Query("select o.envelopeJson from OperationEntity o where o.unitId = :unitId")
    Stream<String> streamEnvelopeJsonByUnitId(@Param("unitId") String unitId);
}
