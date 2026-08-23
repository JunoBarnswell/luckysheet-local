package com.xc.luckysheet.server.persistence;

import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface WorkspaceFolderEntityRepository extends JpaRepository<WorkspaceFolderEntity, String> {
    List<WorkspaceFolderEntity> findBySpaceIdOrderByName(String spaceId);
    List<WorkspaceFolderEntity> findBySpaceIdInOrderByName(java.util.Collection<String> spaceIds);
    Optional<WorkspaceFolderEntity> findByFolderIdAndSpaceId(String folderId, String spaceId);
    List<WorkspaceFolderEntity> findByFolderIdIn(java.util.Collection<String> folderIds);
}
