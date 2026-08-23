package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.CreateSpaceRequest;
import com.xc.luckysheet.server.contract.FolderRequest;
import com.xc.luckysheet.server.contract.FolderResponse;
import com.xc.luckysheet.server.contract.SpaceMemberRequest;
import com.xc.luckysheet.server.contract.SpaceMemberResponse;
import com.xc.luckysheet.server.contract.SpaceResponse;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkspaceSpaceType;
import com.xc.luckysheet.server.persistence.SpaceMemberEntity;
import com.xc.luckysheet.server.persistence.SpaceMemberEntityRepository;
import com.xc.luckysheet.server.persistence.WorkspaceFolderEntity;
import com.xc.luckysheet.server.persistence.WorkspaceFolderEntityRepository;
import com.xc.luckysheet.server.persistence.WorkspaceSpaceEntity;
import com.xc.luckysheet.server.persistence.WorkspaceSpaceEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.Instant;
import java.util.List;
import java.util.UUID;

@Service
public class WorkspaceService {
    private final WorkspaceSpaceEntityRepository spaces;
    private final WorkspaceFolderEntityRepository folders;
    private final SpaceMemberEntityRepository members;
    private final WorkbookEntityRepository workbooks;

    public WorkspaceService(WorkspaceSpaceEntityRepository spaces,
                            WorkspaceFolderEntityRepository folders,
                            SpaceMemberEntityRepository members,
                            WorkbookEntityRepository workbooks) {
        this.spaces = spaces;
        this.folders = folders;
        this.members = members;
        this.workbooks = workbooks;
    }

    @Transactional
    public WorkspaceSpaceEntity ensurePersonalSpace(String subject) {
        if (subject == null || subject.isBlank()) throw ServiceException.unauthenticated("A registered actor is required");
        WorkspaceSpaceEntity existing = spaces.findFirstByOwnerSubjectAndTypeOrderByCreatedAtAsc(subject, WorkspaceSpaceType.PERSONAL).orElse(null);
        if (existing != null) {
            Instant now = Instant.now();
            if (members.findByIdSpaceIdAndIdSubject(existing.getSpaceId(), subject).isEmpty()) {
                members.save(new SpaceMemberEntity(existing.getSpaceId(), subject, WorkbookAclRole.OWNER, now, now));
            }
            return existing;
        }
        Instant now = Instant.now();
        WorkspaceSpaceEntity space = spaces.save(new WorkspaceSpaceEntity(UUID.randomUUID().toString(), "我的云文档",
                WorkspaceSpaceType.PERSONAL, subject, now, now));
        members.save(new SpaceMemberEntity(space.getSpaceId(), subject, WorkbookAclRole.OWNER, now, now));
        return space;
    }

    @Transactional
    public SpaceResponse create(CreateSpaceRequest request, String subject) {
        if (request.kind() != WorkspaceSpaceType.TEAM) {
            throw ServiceException.validation("Personal spaces are created only by the canonical personal-space flow");
        }
        Instant now = Instant.now();
        WorkspaceSpaceEntity space = new WorkspaceSpaceEntity(UUID.randomUUID().toString(), request.name().trim(),
            request.kind(), subject, now, now);
        spaces.save(space);
        members.save(new SpaceMemberEntity(space.getSpaceId(), subject, WorkbookAclRole.OWNER, now, now));
        return response(space, WorkbookAclRole.OWNER);
    }

    @Transactional
    public List<SpaceResponse> list(String subject) {
        ensurePersonalSpace(subject);
        return spaces.findAccessibleTo(subject).stream().map(space -> response(space, effectiveRole(space, subject))).toList();
    }

    public WorkspaceSpaceEntity require(String spaceId, String subject, WorkbookAclRole required) {
        if (spaceId == null || spaceId.isBlank()) throw ServiceException.validation("spaceId is required");
        WorkspaceSpaceEntity space = spaces.findById(spaceId).orElseThrow(() -> ServiceException.notFound("Space not found: " + spaceId));
        WorkbookAclRole role = effectiveRole(space, subject);
        if (role == null || !role.includes(required)) throw ServiceException.forbidden("Space access denied");
        return space;
    }

    public List<FolderResponse> listFolders(String spaceId, String subject) {
        require(spaceId, subject, WorkbookAclRole.VIEWER);
        return folders.findBySpaceIdOrderByName(spaceId).stream().map(this::folderResponse).toList();
    }

    @Transactional
    public FolderResponse createFolder(String spaceId, FolderRequest request, String subject) {
        require(spaceId, subject, WorkbookAclRole.EDITOR);
        if (request.parentFolderId() != null && !request.parentFolderId().isBlank()) {
            WorkspaceFolderEntity parent = folders.findByFolderIdAndSpaceId(request.parentFolderId(), spaceId)
                    .orElseThrow(() -> ServiceException.notFound("Parent folder not found"));
        }
        Instant now = Instant.now();
        WorkspaceFolderEntity folder = new WorkspaceFolderEntity(UUID.randomUUID().toString(), spaceId,
                blankToNull(request.parentFolderId()), request.name().trim(), subject, now, now);
        folders.save(folder);
        return folderResponse(folder);
    }

    @Transactional
    public void deleteFolder(String spaceId, String folderId, String subject) {
        require(spaceId, subject, WorkbookAclRole.EDITOR);
        WorkspaceFolderEntity folder = folders.findByFolderIdAndSpaceId(folderId, spaceId)
                .orElseThrow(() -> ServiceException.notFound("Folder not found"));
        if (folders.existsBySpaceIdAndParentId(spaceId, folderId)) {
            throw ServiceException.conflict("Folder has child folders");
        }
        if (workbooks.existsByFolderId(folderId)) throw ServiceException.conflict("Folder contains workbooks");
        folders.delete(folder);
    }

    @Transactional
    public FolderResponse updateFolder(String folderId, FolderRequest request, String subject) {
        WorkspaceFolderEntity folder = folders.findById(folderId)
                .orElseThrow(() -> ServiceException.notFound("Folder not found"));
        require(folder.getSpaceId(), subject, WorkbookAclRole.EDITOR);
        String parentId = blankToNull(request.parentFolderId());
        if (parentId != null) {
            if (parentId.equals(folderId)) throw ServiceException.validation("Folder cannot contain itself");
            folders.findByFolderIdAndSpaceId(parentId, folder.getSpaceId())
                    .orElseThrow(() -> ServiceException.notFound("Parent folder not found"));
            if (isDescendant(folder.getSpaceId(), folderId, parentId)) {
                throw ServiceException.validation("Folder cannot move below one of its descendants");
            }
        }
        folder.updateLocation(request.name().trim(), parentId, Instant.now());
        folders.save(folder);
        return folderResponse(folder);
    }

    @Transactional
    public void deleteFolder(String folderId, String subject) {
        WorkspaceFolderEntity folder = folders.findById(folderId)
                .orElseThrow(() -> ServiceException.notFound("Folder not found"));
        deleteFolder(folder.getSpaceId(), folderId, subject);
    }

    public List<SpaceMemberResponse> listMembers(String spaceId, String subject) {
        require(spaceId, subject, WorkbookAclRole.VIEWER);
        return members.findByIdSpaceIdOrderByIdSubject(spaceId).stream().map(member ->
                new SpaceMemberResponse(spaceId, member.getId().getSubject(), member.getRole(), member.getCreatedAt(), member.getUpdatedAt())).toList();
    }

    @Transactional
    public SpaceMemberResponse upsertMember(String spaceId, String target, SpaceMemberRequest request, String subject) {
        WorkspaceSpaceEntity space = require(spaceId, subject, WorkbookAclRole.OWNER);
        if (target == null || target.isBlank()) throw ServiceException.validation("Member subject is required");
        if (target.equals(space.getOwnerSubject())) throw ServiceException.validation("The space owner cannot be changed");
        Instant now = Instant.now();
        SpaceMemberEntity member = members.findByIdSpaceIdAndIdSubject(spaceId, target)
                .orElseGet(() -> new SpaceMemberEntity(spaceId, target, request.role(), now, now));
        member.updateRole(request.role(), now);
        members.save(member);
        return new SpaceMemberResponse(spaceId, target, member.getRole(), member.getCreatedAt(), member.getUpdatedAt());
    }

    @Transactional
    public void removeMember(String spaceId, String target, String subject) {
        WorkspaceSpaceEntity space = require(spaceId, subject, WorkbookAclRole.OWNER);
        if (target.equals(space.getOwnerSubject())) throw ServiceException.validation("The space owner cannot be removed");
        members.deleteById(new SpaceMemberEntity.Id(spaceId, target));
    }

    public WorkspaceFolderEntity requireFolder(String spaceId, String folderId, String subject, WorkbookAclRole required) {
        require(spaceId, subject, required);
        if (folderId == null || folderId.isBlank()) return null;
        return folders.findByFolderIdAndSpaceId(folderId, spaceId)
                .orElseThrow(() -> ServiceException.notFound("Folder not found: " + folderId));
    }

    private WorkbookAclRole effectiveRole(WorkspaceSpaceEntity space, String subject) {
        if (space.getOwnerSubject().equals(subject)) return WorkbookAclRole.OWNER;
        return members.findByIdSpaceIdAndIdSubject(space.getSpaceId(), subject).map(SpaceMemberEntity::getRole).orElse(null);
    }

    private SpaceResponse response(WorkspaceSpaceEntity space, WorkbookAclRole role) {
        return new SpaceResponse(space.getSpaceId(), space.getName(), space.getType(), space.getOwnerSubject(), role,
                space.getCreatedAt(), space.getUpdatedAt());
    }

    private FolderResponse folderResponse(WorkspaceFolderEntity folder) {
        return new FolderResponse(folder.getFolderId(), folder.getSpaceId(), folder.getParentId(), folder.getName(),
                folder.getCreatedAt(), folder.getUpdatedAt());
    }

    private static String blankToNull(String value) {
        return value == null || value.isBlank() ? null : value;
    }

    private boolean isDescendant(String spaceId, String ancestorId, String candidateId) {
        java.util.Map<String, WorkspaceFolderEntity> byId = folders.findBySpaceIdOrderByName(spaceId).stream()
                .collect(java.util.stream.Collectors.toMap(WorkspaceFolderEntity::getFolderId, item -> item));
        java.util.Set<String> seen = new java.util.HashSet<>();
        String current = candidateId;
        while (current != null) {
            if (!seen.add(current)) throw ServiceException.conflict("Folder tree contains a cycle");
            if (ancestorId.equals(current)) return true;
            WorkspaceFolderEntity folder = byId.get(current);
            current = folder == null ? null : folder.getParentId();
        }
        return false;
    }
}
