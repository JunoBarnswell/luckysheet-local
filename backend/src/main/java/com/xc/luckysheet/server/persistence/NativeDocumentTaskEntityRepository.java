package com.xc.luckysheet.server.persistence;

import jakarta.persistence.LockModeType;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Lock;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import java.util.Optional;

public interface NativeDocumentTaskEntityRepository extends JpaRepository<NativeDocumentTaskEntity, String> {
    Optional<NativeDocumentTaskEntity> findByTaskIdAndActorSubject(String taskId, String actorSubject);

    @Lock(LockModeType.PESSIMISTIC_WRITE)
    @Query("select t from NativeDocumentTaskEntity t where t.taskId = :id and t.actorSubject = :actor")
    Optional<NativeDocumentTaskEntity> findForUpdate(@Param("id") String id, @Param("actor") String actor);
}
