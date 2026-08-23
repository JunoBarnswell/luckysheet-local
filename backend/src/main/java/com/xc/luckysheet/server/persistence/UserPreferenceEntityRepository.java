package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

public interface UserPreferenceEntityRepository extends JpaRepository<UserPreferenceEntity, String> {
}
