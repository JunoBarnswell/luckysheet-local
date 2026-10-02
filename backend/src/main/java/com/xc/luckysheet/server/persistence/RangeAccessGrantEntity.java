package com.xc.luckysheet.server.persistence;

import com.xc.luckysheet.server.contract.AccessPrincipalKind;
import com.xc.luckysheet.server.contract.RangeAccessLevel;
import jakarta.persistence.Column;
import jakarta.persistence.Embeddable;
import jakarta.persistence.EmbeddedId;
import jakarta.persistence.Entity;
import jakarta.persistence.EnumType;
import jakarta.persistence.Enumerated;
import jakarta.persistence.Table;

import java.io.Serializable;
import java.util.Objects;

@Entity
@Table(name = "workbook_access_grant")
public class RangeAccessGrantEntity {
    @EmbeddedId
    private Id id;

    @Enumerated(EnumType.STRING)
    @Column(name = "access_level", nullable = false, length = 16)
    private RangeAccessLevel access;

    protected RangeAccessGrantEntity() {
    }

    public RangeAccessGrantEntity(String regionId, AccessPrincipalKind principalKind, String principalId,
                                  RangeAccessLevel access) {
        this.id = new Id(regionId, principalKind, principalId);
        this.access = access;
    }

    public Id getId() { return id; }
    public RangeAccessLevel getAccess() { return access; }

    @Embeddable
    public static class Id implements Serializable {
        @Column(name = "region_id", nullable = false, length = 200)
        private String regionId;

        @Enumerated(EnumType.STRING)
        @Column(name = "principal_kind", nullable = false, length = 16)
        private AccessPrincipalKind principalKind;

        @Column(name = "principal_id", nullable = false, length = 500)
        private String principalId;

        protected Id() {
        }

        public Id(String regionId, AccessPrincipalKind principalKind, String principalId) {
            this.regionId = regionId;
            this.principalKind = principalKind;
            this.principalId = principalId;
        }

        public String getRegionId() { return regionId; }
        public AccessPrincipalKind getPrincipalKind() { return principalKind; }
        public String getPrincipalId() { return principalId; }

        @Override
        public boolean equals(Object object) {
            if (this == object) return true;
            if (!(object instanceof Id other)) return false;
            return Objects.equals(regionId, other.regionId)
                    && principalKind == other.principalKind
                    && Objects.equals(principalId, other.principalId);
        }

        @Override
        public int hashCode() { return Objects.hash(regionId, principalKind, principalId); }
    }
}
