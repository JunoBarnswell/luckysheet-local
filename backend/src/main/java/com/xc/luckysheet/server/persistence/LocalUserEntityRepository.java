package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface LocalUserEntityRepository extends JpaRepository<LocalUserEntity, String> {
    Optional<LocalUserEntity> findByUsername(String username);

    List<LocalUserEntity> findAllByOrderByUsernameAsc();
}
