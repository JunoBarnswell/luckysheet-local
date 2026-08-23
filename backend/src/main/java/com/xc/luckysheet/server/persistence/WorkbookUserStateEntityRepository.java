package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

import java.util.Collection;
import java.util.List;
import java.util.Optional;

public interface WorkbookUserStateEntityRepository extends JpaRepository<WorkbookUserStateEntity, WorkbookUserStateEntity.Id> {
    Optional<WorkbookUserStateEntity> findByIdUnitIdAndIdSubject(String unitId, String subject);
    List<WorkbookUserStateEntity> findByIdUnitIdInAndIdSubject(Collection<String> unitIds, String subject);
    void deleteByIdUnitId(String unitId);
}
