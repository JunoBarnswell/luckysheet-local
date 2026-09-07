package com.xc.luckysheet.server.persistence;

import jakarta.persistence.LockModeType;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Lock;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import java.util.Optional;

public interface WorkbookQueryExecutionEntityRepository extends JpaRepository<WorkbookQueryExecutionEntity, String> {
    @Lock(LockModeType.PESSIMISTIC_WRITE)
    @Query("select q from WorkbookQueryExecutionEntity q where q.unitId = :unitId and q.queryId = :queryId")
    Optional<WorkbookQueryExecutionEntity> findForUpdate(@Param("unitId") String unitId, @Param("queryId") String queryId);
}
