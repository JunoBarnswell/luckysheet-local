package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.Collection;
import java.util.List;
import java.util.Optional;

public interface SpaceMemberEntityRepository extends JpaRepository<SpaceMemberEntity, SpaceMemberEntity.Id> {
    Optional<SpaceMemberEntity> findByIdSpaceIdAndIdSubject(String spaceId, String subject);
    List<SpaceMemberEntity> findByIdSpaceIdOrderByIdSubject(String spaceId);
    List<SpaceMemberEntity> findByIdSpaceIdInAndIdSubject(Collection<String> spaceIds, String subject);

    @Query("select m from SpaceMemberEntity m where m.id.spaceId in :spaceIds")
    List<SpaceMemberEntity> findForSpaces(@Param("spaceIds") Collection<String> spaceIds);
}
