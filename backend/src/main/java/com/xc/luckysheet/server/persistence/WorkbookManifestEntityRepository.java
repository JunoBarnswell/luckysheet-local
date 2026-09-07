package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import java.util.Optional;

public interface WorkbookManifestEntityRepository extends JpaRepository<WorkbookManifestEntity, String> {
    Optional<WorkbookManifestEntity> findByUnitIdAndRevision(String unitId, long revision);
    void deleteByUnitId(String unitId);
}
