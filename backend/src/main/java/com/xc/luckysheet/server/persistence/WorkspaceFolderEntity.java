package com.xc.luckysheet.server.persistence;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;

import java.time.Instant;

@Entity
@Table(name = "workspace_folder", indexes = @Index(name = "workspace_folder_space_parent_idx", columnList = "space_id,parent_id,name"))
public class WorkspaceFolderEntity {
    @Id
    @Column(name = "folder_id", nullable = false, length = 200)
    private String folderId;

    @Column(name = "space_id", nullable = false, length = 200)
    private String spaceId;

    @Column(name = "parent_id", length = 200)
    private String parentId;

    @Column(name = "name", nullable = false, length = 200)
    private String name;

    @Column(name = "created_by", nullable = false, length = 500)
    private String createdBy;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected WorkspaceFolderEntity() {}

    public WorkspaceFolderEntity(String folderId, String spaceId, String parentId, String name, String createdBy,
                                 Instant createdAt, Instant updatedAt) {
        this.folderId = folderId;
        this.spaceId = spaceId;
        this.parentId = parentId;
        this.name = name;
        this.createdBy = createdBy;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
    }

    public String getFolderId() { return folderId; }
    public String getSpaceId() { return spaceId; }
    public String getParentId() { return parentId; }
    public String getName() { return name; }
    public String getCreatedBy() { return createdBy; }
    public Instant getCreatedAt() { return createdAt; }
    public Instant getUpdatedAt() { return updatedAt; }

    public void rename(String name, Instant updatedAt) {
        this.name = name;
        this.updatedAt = updatedAt;
    }

    public void updateLocation(String name, String parentId, Instant updatedAt) {
        this.name = name;
        this.parentId = parentId;
        this.updatedAt = updatedAt;
    }
}
