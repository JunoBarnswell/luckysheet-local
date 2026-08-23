package com.xc.luckysheet.server.persistence;

import com.xc.luckysheet.server.contract.WorkbookAclRole;
import jakarta.persistence.Column;
import jakarta.persistence.Embeddable;
import jakarta.persistence.EmbeddedId;
import jakarta.persistence.Entity;
import jakarta.persistence.EnumType;
import jakarta.persistence.Enumerated;
import jakarta.persistence.Index;
import jakarta.persistence.Table;

import java.io.Serializable;
import java.time.Instant;
import java.util.Objects;

@Entity
@Table(name = "space_member", indexes = @Index(name = "space_member_subject_idx", columnList = "subject,space_id"))
public class SpaceMemberEntity {
    @EmbeddedId
    private Id id;

    @Enumerated(EnumType.STRING)
    @Column(name = "role", nullable = false, length = 16)
    private WorkbookAclRole role;

    @Column(name = "created_at", nullable = false)
    private Instant createdAt;

    @Column(name = "updated_at", nullable = false)
    private Instant updatedAt;

    protected SpaceMemberEntity() {}

    public SpaceMemberEntity(String spaceId, String subject, WorkbookAclRole role, Instant createdAt, Instant updatedAt) {
        this.id = new Id(spaceId, subject);
        this.role = role;
        this.createdAt = createdAt;
        this.updatedAt = updatedAt;
    }

    public Id getId() { return id; }
    public WorkbookAclRole getRole() { return role; }
    public Instant getCreatedAt() { return createdAt; }
    public Instant getUpdatedAt() { return updatedAt; }

    public void updateRole(WorkbookAclRole role, Instant updatedAt) {
        this.role = role;
        this.updatedAt = updatedAt;
    }

    @Embeddable
    public static class Id implements Serializable {
        @Column(name = "space_id", nullable = false, length = 200)
        private String spaceId;

        @Column(name = "subject", nullable = false, length = 500)
        private String subject;

        protected Id() {}

        public Id(String spaceId, String subject) {
            this.spaceId = spaceId;
            this.subject = subject;
        }

        public String getSpaceId() { return spaceId; }
        public String getSubject() { return subject; }

        @Override
        public boolean equals(Object object) {
            if (this == object) return true;
            if (!(object instanceof Id other)) return false;
            return Objects.equals(spaceId, other.spaceId) && Objects.equals(subject, other.subject);
        }

        @Override
        public int hashCode() { return Objects.hash(spaceId, subject); }
    }
}
