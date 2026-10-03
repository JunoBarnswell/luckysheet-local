package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.contract.AccessPrincipalKind;
import com.xc.luckysheet.server.contract.EffectiveAccessRegion;
import com.xc.luckysheet.server.contract.RangeAccessGrant;
import com.xc.luckysheet.server.contract.RangeAccessLevel;
import com.xc.luckysheet.server.contract.RangeAccessRegion;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookRole;

import java.util.ArrayList;
import java.util.List;

/** One resolver for point, rectangle, and multi-range access decisions. */
public final class RangeAccessResolver {
    private final RangeAccessIndex index;
    private final List<RangeAccessRegion> regions;
    private final RangeAccessContext context;

    public RangeAccessResolver(RangeAccessIndex index, List<RangeAccessRegion> regions, RangeAccessContext context) {
        this.index = index;
        this.regions = List.copyOf(regions);
        this.context = context;
    }

    public long accessRevision() {
        return context.accessRevision();
    }

    public RangeAccessLevel effectiveAccess(String sheetId, int row, int column) {
        return effectiveAccess(new RangeRef(sheetId, row, row, column, column));
    }

    public RangeAccessLevel effectiveAccess(RangeRef range) {
        List<RangeAccessRegion> matches = index.intersecting(range);
        RangeAccessLevel roleCeiling = roleCeiling(context.workbookRole());
        if (matches.isEmpty()) return roleCeiling;
        RangeAccessLevel effective = roleCeiling;
        for (RangeAccessRegion region : matches) {
            RangeAccessLevel regionAccess = restrict(resolveRegion(region), roleCeiling);
            if (regionAccess.ordinal() < effective.ordinal()) effective = regionAccess;
        }
        return effective;
    }

    public boolean canRead(RangeRef range) {
        if (context.workbookRole() == WorkbookRole.OWNER) return true;
        for (RangeAccessRegion region : index.intersecting(range)) {
            if (restrict(resolveRegion(region), roleCeiling(context.workbookRole())) == RangeAccessLevel.HIDDEN) return false;
        }
        return true;
    }

    public boolean canEdit(RangeRef range) {
        if (roleCeiling(context.workbookRole()) != RangeAccessLevel.EDIT) return false;
        for (RangeAccessRegion region : index.intersecting(range)) {
            if (restrict(resolveRegion(region), RangeAccessLevel.EDIT) != RangeAccessLevel.EDIT) return false;
        }
        return true;
    }

    public void requireCanRead(List<RangeRef> ranges) {
        for (RangeRef range : ranges) {
            if (!canRead(range)) throw new ServiceException("ACCESS_HIDDEN", 403, "Requested range contains data the current subject cannot view");
        }
    }

    public void requireCanEdit(List<RangeRef> ranges) {
        if (roleCeiling(context.workbookRole()) != RangeAccessLevel.EDIT) {
            throw new ServiceException("ACCESS_DENIED", 403, "Workbook role does not permit editing");
        }
        for (RangeRef range : ranges) {
            if (!canEdit(range)) throw new ServiceException("ACCESS_DENIED", 403, "Mutation affects a range the current subject cannot edit");
        }
    }

    public List<RangeAccessRegion> hiddenRegions() {
        if (context.workbookRole() == WorkbookRole.OWNER) return List.of();
        return regions.stream().filter(region -> resolveRegion(region) == RangeAccessLevel.HIDDEN
                        || restrict(resolveRegion(region), roleCeiling(context.workbookRole())) == RangeAccessLevel.HIDDEN)
                .toList();
    }

    public List<EffectiveAccessRegion> effectiveRegions() {
        return regions.stream().map(region -> new EffectiveAccessRegion(region.range(),
                restrict(resolveRegion(region), roleCeiling(context.workbookRole())))).toList();
    }

    private RangeAccessLevel resolveRegion(RangeAccessRegion region) {
        if (context.workbookRole() == WorkbookRole.OWNER) return RangeAccessLevel.EDIT;
        RangeAccessLevel subject = null;
        List<RangeAccessLevel> groups = new ArrayList<>();
        RangeAccessLevel everyone = null;
        for (RangeAccessGrant grant : region.grants()) {
            switch (grant.principal().kind()) {
                case SUBJECT -> {
                    if (context.subject().equals(grant.principal().id())) subject = grant.access();
                }
                case GROUP -> {
                    if (context.groups().contains(grant.principal().id())) groups.add(grant.access());
                }
                case EVERYONE -> everyone = grant.access();
            }
        }
        if (subject != null) return subject;
        if (!groups.isEmpty()) return groups.stream().max(java.util.Comparator.comparingInt(Enum::ordinal)).orElseThrow();
        return everyone == null ? region.defaultAccess() : everyone;
    }

    private static RangeAccessLevel roleCeiling(WorkbookRole role) {
        return role == WorkbookRole.OWNER || role == WorkbookRole.EDITOR
                ? RangeAccessLevel.EDIT : RangeAccessLevel.READ;
    }

    private static RangeAccessLevel restrict(RangeAccessLevel access, RangeAccessLevel ceiling) {
        return access.ordinal() <= ceiling.ordinal() ? access : ceiling;
    }
}
