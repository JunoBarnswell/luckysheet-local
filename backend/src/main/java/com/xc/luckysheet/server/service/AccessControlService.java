package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.AclEntry;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;

import java.time.Instant;
import java.util.List;

@Service
public class AccessControlService {
    private final WorkbookStore store;
    private final GuestShareService guestShares;

    public AccessControlService(WorkbookStore store, GuestShareService guestShares) {
        this.store = store;
        this.guestShares = guestShares;
    }

    public WorkbookAclRole require(String unitId, String subject, WorkbookAclRole required) {
        WorkbookAclRole actual = store.findRole(unitId, subject).orElse(null);
        if (actual == null) actual = guestShares.roleFor(unitId, subject);
        if (actual == null) throw ServiceException.forbidden("Workbook access denied");
        if (!actual.includes(required)) throw ServiceException.forbidden("Workbook role " + required + " is required");
        return actual;
    }

    public List<AclEntry> list(String unitId, String subject) {
        require(unitId, subject, WorkbookAclRole.OWNER);
        return store.listAcl(unitId);
    }

    public AclEntry grant(String unitId, String actor, String target, WorkbookAclRole role) {
        require(unitId, actor, WorkbookAclRole.OWNER);
        if (target == null || target.isBlank()) throw ServiceException.validation("ACL subject is required");
        if (role == null || role == WorkbookAclRole.OWNER) throw ServiceException.validation("Only editor, commenter or viewer may be granted");
        Instant now = Instant.now();
        store.upsertAcl(unitId, target, role, now);
        return store.listAcl(unitId).stream().filter(entry -> entry.subject().equals(target)).findFirst()
                .orElseThrow(() -> new IllegalStateException("ACL write was not persisted"));
    }

    public void revoke(String unitId, String actor, String target) {
        require(unitId, actor, WorkbookAclRole.OWNER);
        if (store.findRole(unitId, target).orElse(null) == WorkbookAclRole.OWNER) {
            throw ServiceException.forbidden("The workbook owner cannot be revoked");
        }
        store.deleteAcl(unitId, target);
    }
}
