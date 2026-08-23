package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Lock;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import jakarta.persistence.LockModeType;
import java.util.Collection;
import java.util.List;
import java.util.Optional;
import org.springframework.data.domain.Pageable;

public interface WorkbookEntityRepository extends JpaRepository<WorkbookEntity, String> {
    @Lock(LockModeType.PESSIMISTIC_WRITE)
    @Query("select w from WorkbookEntity w where w.unitId = :unitId")
    Optional<WorkbookEntity> findForUpdate(@Param("unitId") String unitId);

    /** One catalog query for the actor; user state is batch-loaded afterwards. */
    @Query("""
            select distinct w from WorkbookEntity w
            left join SpaceMemberEntity sm on sm.id.spaceId = w.spaceId and sm.id.subject = :subject
            where (w.ownerSubject = :subject
                or exists (select a from WorkbookAclEntity a where a.id.unitId = w.unitId and a.id.subject = :subject)
                or sm.id.subject = :subject)
              and (:includeTrash = true or w.deletedAt is null)
              and (:trashOnly = false or w.deletedAt is not null)
              and (:sharedOnly = false or w.ownerSubject <> :subject)
              and (:ownedOnly = false or w.ownerSubject = :subject)
              and (:spaceId is null or w.spaceId = :spaceId)
              and (:folderId is null or w.folderId = :folderId)
              and (:query is null or lower(w.name) like lower(concat('%', :query, '%')))
            order by w.updatedAt desc
            """)
    List<WorkbookEntity> findCatalogCandidates(
            @Param("subject") String subject,
            @Param("includeTrash") boolean includeTrash,
            @Param("trashOnly") boolean trashOnly,
            @Param("sharedOnly") boolean sharedOnly,
            @Param("ownedOnly") boolean ownedOnly,
            @Param("spaceId") String spaceId,
            @Param("folderId") String folderId,
            @Param("query") String query,
            Pageable pageable
    );

    boolean existsByFolderId(String folderId);
}
