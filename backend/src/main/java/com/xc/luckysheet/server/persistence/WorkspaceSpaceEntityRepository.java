package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.List;
import java.util.Optional;

public interface WorkspaceSpaceEntityRepository extends JpaRepository<WorkspaceSpaceEntity, String> {
    Optional<WorkspaceSpaceEntity> findByOwnerSubjectAndType(String ownerSubject, com.xc.luckysheet.server.contract.WorkspaceSpaceType type);

    @Query("select distinct s from WorkspaceSpaceEntity s left join SpaceMemberEntity m on m.id.spaceId = s.spaceId and m.id.subject = :subject where s.ownerSubject = :subject or m.id.subject = :subject order by s.updatedAt desc")
    List<WorkspaceSpaceEntity> findAccessibleTo(@Param("subject") String subject);
}
