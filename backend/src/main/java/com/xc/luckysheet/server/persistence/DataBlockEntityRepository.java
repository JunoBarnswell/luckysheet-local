package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

public interface DataBlockEntityRepository extends JpaRepository<DataBlockEntity, DataBlockEntity.Id> {
}
