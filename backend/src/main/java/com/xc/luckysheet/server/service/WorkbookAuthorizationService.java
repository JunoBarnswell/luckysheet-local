package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.persistence.SpaceMemberEntity;
import com.xc.luckysheet.server.persistence.SpaceMemberEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookAclEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import org.springframework.stereotype.Service;

import java.util.Optional;

/** Calculates effective access from workbook ACL, owner and space membership. */
@Service
public class WorkbookAuthorizationService {
    private final WorkbookEntityRepository workbooks;
    private final WorkbookAclEntityRepository acl;
    private final SpaceMemberEntityRepository members;

    public WorkbookAuthorizationService(WorkbookEntityRepository workbooks,
                                         WorkbookAclEntityRepository acl,
                                         SpaceMemberEntityRepository members) {
        this.workbooks = workbooks;
        this.acl = acl;
        this.members = members;
    }

    public Optional<WorkbookAclRole> role(String unitId, String subject) {
        WorkbookEntity workbook = workbooks.findById(unitId).orElse(null);
        if (workbook == null) return Optional.empty();
        WorkbookAclRole effective = null;
        if (subject.equals(workbook.getOwnerSubject())) effective = WorkbookAclRole.OWNER;
        WorkbookAclRole direct = acl.findForSubject(unitId, subject).map(e -> e.getRole()).orElse(null);
        effective = max(effective, direct);
        if (workbook.getSpaceId() != null && !workbook.getSpaceId().isBlank()) {
            WorkbookAclRole spaceRole = members.findByIdSpaceIdAndIdSubject(workbook.getSpaceId(), subject)
                    .map(SpaceMemberEntity::getRole).orElse(null);
            effective = max(effective, spaceRole);
        }
        return Optional.ofNullable(effective);
    }

    private WorkbookAclRole max(WorkbookAclRole left, WorkbookAclRole right) {
        if (left == null) return right;
        if (right == null) return left;
        return left.includes(right) ? left : right;
    }
}
