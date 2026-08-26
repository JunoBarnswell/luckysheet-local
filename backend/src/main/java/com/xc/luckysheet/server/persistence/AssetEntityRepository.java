package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface AssetEntityRepository extends JpaRepository<AssetEntity, AssetEntity.Id> {
    Optional<AssetEntity> findByIdUnitIdAndContentHash(String unitId, String contentHash);
    List<AssetEntity> findAllByIdUnitId(String unitId);
}
