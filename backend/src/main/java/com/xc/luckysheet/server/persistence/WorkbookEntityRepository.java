package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Lock;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import jakarta.persistence.LockModeType;
import java.util.List;
import java.util.Optional;

public interface WorkbookEntityRepository extends JpaRepository<WorkbookEntity, String> {
    @Lock(LockModeType.PESSIMISTIC_WRITE)
    @Query("select w from WorkbookEntity w where w.unitId = :unitId")
    Optional<WorkbookEntity> findForUpdate(@Param("unitId") String unitId);

    @Query("select w from WorkbookEntity w where exists (select a from WorkbookAclEntity a where a.id.unitId = w.unitId and a.id.subject = :subject) order by w.updatedAt desc")
    List<WorkbookEntity> findAllAccessibleTo(@Param("subject") String subject);
}
