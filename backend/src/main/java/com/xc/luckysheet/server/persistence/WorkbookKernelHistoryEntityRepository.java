package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import java.util.Optional;

public interface WorkbookKernelHistoryEntityRepository extends JpaRepository<WorkbookKernelHistoryEntity, String> {
    Optional<WorkbookKernelHistoryEntity> findByUnitIdAndOperationId(String unitId, String operationId);
    void deleteByUnitId(String unitId);
}
