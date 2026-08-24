package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.AclEntry;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.Instant;
import java.util.List;

@Service
public class AccessControlService {
    private final WorkbookStore store;
    private final GuestShareService guestShares;
    private final WorkbookAuthorizationService authorization;

    public AccessControlService(WorkbookStore store, GuestShareService guestShares, WorkbookAuthorizationService authorization) {
        this.store = store;
        this.guestShares = guestShares;
        this.authorization = authorization;
    }

    public WorkbookAclRole require(String unitId, String subject, WorkbookAclRole required) {
        WorkbookAclRole actual = authorization.role(unitId, subject).orElse(null);
        if (actual == null) actual = guestShares.roleFor(unitId, subject);
        if (actual == null) throw ServiceException.forbidden("Workbook access denied");
        if (!actual.includes(required)) throw ServiceException.forbidden("Workbook role " + required + " is required");
        return actual;
    }

    /** Returns only a role derived from persistent ACL or a verified share token. */
    public WorkbookAclRole currentRole(String unitId, String subject) {
        return require(unitId, subject, WorkbookAclRole.VIEWER);
    }

    public List<AclEntry> list(String unitId, String subject) {
        require(unitId, subject, WorkbookAclRole.OWNER);
        return store.listAcl(unitId);
    }

    @Transactional
    public AclEntry grant(String unitId, String actor, String target, WorkbookAclRole role) {
        store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found"));
        require(unitId, actor, WorkbookAclRole.OWNER);
        if (target == null || target.isBlank()) throw ServiceException.validation("ACL subject is required");
        if (role == null || role == WorkbookAclRole.OWNER) throw ServiceException.validation("Only editor, commenter or viewer may be granted");
        Instant now = Instant.now();
        store.upsertAcl(unitId, target, role, now);
        return store.listAcl(unitId).stream().filter(entry -> entry.subject().equals(target)).findFirst()
                .orElseThrow(() -> new IllegalStateException("ACL write was not persisted"));
    }

    @Transactional
    public void revoke(String unitId, String actor, String target) {
        store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found"));
        require(unitId, actor, WorkbookAclRole.OWNER);
        if (authorization.role(unitId, target).orElse(null) == WorkbookAclRole.OWNER) {
            throw ServiceException.forbidden("The workbook owner cannot be revoked");
        }
        store.deleteAcl(unitId, target);
    }
}
