package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.util.List;
import java.util.Optional;

public interface WorkbookAclEntityRepository extends JpaRepository<WorkbookAclEntity, WorkbookAclEntity.Id> {
    @Query("select a from WorkbookAclEntity a where a.id.unitId = :unitId order by a.id.subject")
    List<WorkbookAclEntity> findAllForWorkbook(@Param("unitId") String unitId);

    @Query("select a from WorkbookAclEntity a where a.id.unitId = :unitId and a.id.subject = :subject")
    Optional<WorkbookAclEntity> findForSubject(@Param("unitId") String unitId, @Param("subject") String subject);
}
