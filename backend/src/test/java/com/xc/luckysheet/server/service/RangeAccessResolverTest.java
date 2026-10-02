package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.AccessPrincipal;
import com.xc.luckysheet.server.contract.AccessPrincipalKind;
import com.xc.luckysheet.server.contract.RangeAccessGrant;
import com.xc.luckysheet.server.contract.RangeAccessLevel;
import com.xc.luckysheet.server.contract.RangeAccessRegion;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RangeAccessResolverTest {
    private static final String SHEET = "sheet-1";

    @Test
    void subjectGrantOverridesGroupAndEveryoneAndWorkbookRoleCapsTheResult() {
        RangeAccessRegion region = region(RangeAccessLevel.HIDDEN,
                grant(AccessPrincipalKind.EVERYONE, null, RangeAccessLevel.HIDDEN),
                grant(AccessPrincipalKind.GROUP, "analysts", RangeAccessLevel.EDIT),
                grant(AccessPrincipalKind.SUBJECT, "alice", RangeAccessLevel.READ));

        RangeAccessResolver editor = resolver(WorkbookAclRole.EDITOR, List.of("analysts"), region);
        assertEquals(RangeAccessLevel.READ, editor.effectiveAccess(SHEET, 2, 2));
        assertTrue(editor.canRead(new RangeRef(SHEET, 2, 2, 2, 2)));
        assertFalse(editor.canEdit(new RangeRef(SHEET, 2, 2, 2, 2)));

        RangeAccessResolver viewer = resolver(WorkbookAclRole.VIEWER, List.of("analysts"), region);
        assertEquals(RangeAccessLevel.READ, viewer.effectiveAccess(SHEET, 2, 2));
        assertFalse(viewer.canEdit(new RangeRef(SHEET, 2, 2, 2, 2)));

        RangeAccessResolver groupMember = new RangeAccessResolver(new RangeAccessIndex(List.of(region)), List.of(region),
                new RangeAccessContext("bob", WorkbookAclRole.EDITOR, List.of("analysts"), 12));
        assertEquals(RangeAccessLevel.EDIT, groupMember.effectiveAccess(SHEET, 2, 2));
        assertTrue(groupMember.canEdit(new RangeRef(SHEET, 2, 2, 2, 2)));
    }

    @Test
    void hiddenIntersectionRejectsWholeMutationAndOwnerRetainsAccess() {
        RangeAccessRegion region = region(RangeAccessLevel.HIDDEN);
        RangeRef partiallyHidden = new RangeRef(SHEET, 4, 6, 3, 5);
        RangeAccessResolver editor = resolver(WorkbookAclRole.EDITOR, List.of(), region);

        assertFalse(editor.canRead(partiallyHidden));
        assertFalse(editor.canEdit(partiallyHidden));
        assertThrows(ServiceException.class, () -> editor.requireCanRead(List.of(partiallyHidden)));
        assertThrows(ServiceException.class, () -> editor.requireCanEdit(List.of(partiallyHidden)));

        RangeAccessResolver owner = resolver(WorkbookAclRole.OWNER, List.of(), region);
        assertTrue(owner.canRead(partiallyHidden));
        assertTrue(owner.canEdit(partiallyHidden));
        assertTrue(owner.hiddenRegions().isEmpty());
    }

    private static RangeAccessResolver resolver(WorkbookAclRole role, List<String> groups, RangeAccessRegion... regions) {
        List<RangeAccessRegion> configured = List.of(regions);
        return new RangeAccessResolver(new RangeAccessIndex(configured), configured,
                new RangeAccessContext("alice", role, groups, 12));
    }

    private static RangeAccessRegion region(RangeAccessLevel defaultAccess, RangeAccessGrant... grants) {
        return new RangeAccessRegion("region-1", "book-1", SHEET,
                new RangeRef(SHEET, 2, 5, 2, 5), defaultAccess, List.of(grants), "owner", Instant.EPOCH, Instant.EPOCH);
    }

    private static RangeAccessGrant grant(AccessPrincipalKind kind, String id, RangeAccessLevel access) {
        return new RangeAccessGrant(new AccessPrincipal(kind, id), access);
    }
}
