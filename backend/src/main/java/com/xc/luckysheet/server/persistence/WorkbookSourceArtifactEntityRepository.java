package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

public interface WorkbookSourceArtifactEntityRepository extends JpaRepository<WorkbookSourceArtifactEntity, String> {
    java.util.List<WorkbookSourceArtifactEntity> findByUnitIdIn(java.util.Collection<String> unitIds);
}
