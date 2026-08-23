package com.xc.luckysheet.server.persistence;

import com.xc.luckysheet.server.contract.WorkspaceSpaceType;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.EnumType;
import jakarta.persistence.Enumerated;
import jakarta.persistence.Id;
import jakarta.persistence.Index;
import jakarta.persistence.Table;

import java.time.Instant;

@Entity
@Table(name = "workspace_space", indexes = @Index(name = "workspace_space_owner_idx", columnList = "owner_subject,updated_at"))
public class WorkspaceSpaceEntity {
    @Id
    @Column(name = "space_id", nullable = false, length = 200)
    private String spaceId;

    @Column(name = "name", nullable = false, length = 200)
    private String name;

    @Enumerated(EnumType.STRING)
    @Column(name = "type", nullable = false, length = 16)
    private WorkspaceSpaceType type;

    @Column(name = "owner_subject", nullable = false, length = 500)
    private String ownerSubject;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected WorkspaceSpaceEntity() {}

    public WorkspaceSpaceEntity(String spaceId, String name, WorkspaceSpaceType type, String ownerSubject,
                                Instant createdAt, Instant updatedAt) {
        this.spaceId = spaceId;
        this.name = name;
        this.type = type;
        this.ownerSubject = ownerSubject;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
    }

    public String getSpaceId() { return spaceId; }
    public String getName() { return name; }
    public WorkspaceSpaceType getType() { return type; }
    public String getOwnerSubject() { return ownerSubject; }
    public Instant getCreatedAt() { return createdAt; }
    public Instant getUpdatedAt() { return updatedAt; }

    public void rename(String name, Instant updatedAt) {
        this.name = name;
        this.updatedAt = updatedAt;
    }
}
