package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.contract.WorkbookSource;
import com.xc.luckysheet.server.contract.WorkbookStorageLocation;
import com.xc.luckysheet.server.persistence.SpaceMemberEntity;
import com.xc.luckysheet.server.persistence.SpaceMemberEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookAclEntity;
import com.xc.luckysheet.server.persistence.WorkbookAclEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class WorkbookAuthorizationServiceTest {
    @Test
    void effectiveRoleUsesTheStrongestDirectAclAndSpaceMembership() {
        WorkbookEntityRepository workbooks = mock(WorkbookEntityRepository.class);
        WorkbookAclEntityRepository acl = mock(WorkbookAclEntityRepository.class);
        SpaceMemberEntityRepository members = mock(SpaceMemberEntityRepository.class);
        WorkbookAuthorizationService service = new WorkbookAuthorizationService(workbooks, acl, members);
        Instant now = Instant.now();
        when(workbooks.findById("book-1")).thenReturn(Optional.of(new WorkbookEntity("book-1", "Book", "{}", 0, 0,
                now, now, "owner", "space-1", null, WorkbookStorageLocation.REMOTE, WorkbookSource.NATIVE,
                WorkbookLifecycle.ACTIVE, null)));
        when(acl.findForSubject("book-1", "member")).thenReturn(Optional.of(
                new WorkbookAclEntity("book-1", "member", WorkbookAclRole.VIEWER, now, now)));
        when(members.findByIdSpaceIdAndIdSubject("space-1", "member")).thenReturn(Optional.of(
                new SpaceMemberEntity("space-1", "member", WorkbookAclRole.EDITOR, now, now)));

        assertEquals(Optional.of(WorkbookAclRole.EDITOR), service.role("book-1", "member"));
        assertTrue(service.role("missing", "member").isEmpty());
    }
}
