package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.AccessPrincipal;
import com.xc.luckysheet.server.contract.AccessPrincipalKind;
import com.xc.luckysheet.server.contract.EffectiveAccessRegion;
import com.xc.luckysheet.server.contract.RangeAccessChangeResponse;
import com.xc.luckysheet.server.contract.RangeAccessGrant;
import com.xc.luckysheet.server.contract.RangeAccessLevel;
import com.xc.luckysheet.server.contract.RangeAccessRegion;
import com.xc.luckysheet.server.contract.RangeAccessRegionRequest;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAccessProjection;
import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.mutation.MutationDescriptorRegistry;
import com.xc.luckysheet.server.persistence.RangeAccessGrantEntity;
import com.xc.luckysheet.server.persistence.RangeAccessRegionEntity;
import com.xc.luckysheet.server.store.WorkbookRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.Instant;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;

@Service
public class RangeAccessService {
    private static final int MAX_ROW_INDEX = 1_048_575;
    private static final int MAX_COLUMN_INDEX = 16_383;

    private final WorkbookStore store;
    private final AccessControlService access;
    private final ObjectMapper mapper;
    private final MutationDescriptorRegistry structuralTransform;
    private final Map<IndexKey, AccessIndexSnapshot> indexes = new ConcurrentHashMap<>();

    public RangeAccessService(WorkbookStore store, AccessControlService access, ObjectMapper mapper,
                              MutationDescriptorRegistry structuralTransform) {
        this.store = store;
        this.access = access;
        this.mapper = mapper;
        this.structuralTransform = structuralTransform;
    }

    public WorkbookAccessProjection projection(String unitId, String subject, Collection<String> groups) {
        WorkbookRole role = access.currentRole(unitId, subject);
        RangeAccessResolver resolver = resolver(unitId, subject, role, groups);
        return new WorkbookAccessProjection(unitId, role, resolver.accessRevision(), resolver.effectiveRegions());
    }

    @Transactional(readOnly = true)
    public List<RangeAccessRegion> list(String unitId, String subject) {
        access.require(unitId, subject, WorkbookRole.OWNER);
        return loadRegions(unitId);
    }

    @Transactional
    public RangeAccessChangeResponse create(String unitId, String subject, RangeAccessRegionRequest request) {
        requireOwnerAndLockWorkbook(unitId, subject);
        validateRequest(unitId, request);
        RangeRef range = request.range();
        requireNoOverlap(unitId, range, null);
        Instant now = Instant.now();
        RangeAccessRegionEntity entity = new RangeAccessRegionEntity(UUID.randomUUID().toString(), unitId,
                request.sheetId(), range, request.defaultAccess(), subject, now);
        store.saveAccessRegion(entity, grantEntities(entity.getId(), request.grants()));
        long revision = store.incrementAccessRevision(unitId);
        evict(unitId);
        return new RangeAccessChangeResponse(toRegion(entity, request.grants()), revision);
    }

    @Transactional
    public RangeAccessChangeResponse update(String unitId, String regionId, String subject, RangeAccessRegionRequest request) {
        requireOwnerAndLockWorkbook(unitId, subject);
        RangeAccessRegionEntity entity = store.findAccessRegion(unitId, regionId)
                .orElseThrow(() -> new ServiceException("ACCESS_REGION_NOT_FOUND", 404, "Range access region was not found"));
        validateRequest(unitId, request);
        requireNoOverlap(unitId, request.range(), regionId);
        entity.update(request.sheetId(), request.range(), request.defaultAccess(), Instant.now());
        store.saveAccessRegion(entity, grantEntities(entity.getId(), request.grants()));
        long revision = store.incrementAccessRevision(unitId);
        evict(unitId);
        return new RangeAccessChangeResponse(toRegion(entity, request.grants()), revision);
    }

    @Transactional
    public long delete(String unitId, String regionId, String subject) {
        requireOwnerAndLockWorkbook(unitId, subject);
        RangeAccessRegionEntity entity = store.findAccessRegion(unitId, regionId)
                .orElseThrow(() -> new ServiceException("ACCESS_REGION_NOT_FOUND", 404, "Range access region was not found"));
        store.deleteAccessRegion(entity);
        long revision = store.incrementAccessRevision(unitId);
        evict(unitId);
        return revision;
    }

    public RangeAccessResolver resolver(String unitId, String subject, Collection<String> groups) {
        WorkbookRole role = access.currentRole(unitId, subject);
        return resolver(unitId, subject, role, groups);
    }

    public RangeAccessResolver resolver(String unitId, String subject, WorkbookRole role, Collection<String> groups) {
        long revision = store.accessRevision(unitId);
        IndexKey key = new IndexKey(unitId, revision);
        AccessIndexSnapshot snapshot = indexes.computeIfAbsent(key, ignored -> loadIndex(unitId, revision));
        RangeAccessContext context = new RangeAccessContext(subject, role, groups == null ? List.of() : List.copyOf(groups), revision);
        return new RangeAccessResolver(snapshot.index(), snapshot.regions(), context);
    }

    /** Apply ACL geometry/lifecycle changes from the canonical server structural mutation. */
    @Transactional
    public Long applyStructuralMutation(String unitId, OperationMutation mutation, JsonNode snapshotAfter) {
        boolean changed = switch (mutation.id()) {
            case "rows.inserted", "rows.deleted", "columns.inserted", "columns.deleted" ->
                    transformAxis(unitId, mutation.sheetId(), mutation.params().path("at").asInt(-1),
                            mutation.params().path("count").asInt(-1), mutation.id().startsWith("rows."), mutation.id().endsWith("inserted"));
            case "sheet.remove" -> deleteSheet(unitId, mutation.params().path("id").asText());
            case "sheet.duplicated" -> duplicateSheet(unitId, mutation.params().path("sourceSheetId").asText(),
                    mutation.params().path("newId").asText(), snapshotAfter);
            default -> false;
        };
        if (!changed) return null;
        long revision = store.incrementAccessRevision(unitId);
        evict(unitId);
        return revision;
    }

    @Transactional
    public long bumpAccessRevision(String unitId) {
        long revision = store.incrementAccessRevision(unitId);
        evict(unitId);
        return revision;
    }

    private boolean transformAxis(String unitId, String sheetId, int at, int count, boolean rows, boolean insert) {
        if (sheetId == null || sheetId.isBlank() || at < 0 || count < 1) {
            throw ServiceException.validation("Structural range access transform is invalid");
        }
        boolean changed = false;
        for (RangeAccessRegionEntity entity : regionEntities(unitId)) {
            if (!sheetId.equals(entity.getSheetId())) continue;
            RangeAccessRegion current = mapRegions(List.of(entity), store.listAccessGrants(List.of(entity.getId()))).getFirst();
            RangeRef mapped = structuralTransform.transformAccessRange(current.range(), sheetId, rows, at, count, insert);
            if (java.util.Objects.equals(mapped, current.range())) continue;
            changed = true;
            if (mapped == null) store.deleteAccessRegion(entity);
            else {
                entity.update(sheetId, mapped, entity.getDefaultAccess(), Instant.now());
                store.saveAccessRegion(entity, grantEntities(entity.getId(), current.grants()));
            }
        }
        return changed;
    }

    private boolean deleteSheet(String unitId, String sheetId) {
        if (sheetId == null || sheetId.isBlank()) throw ServiceException.validation("Removed sheet identity is invalid");
        List<RangeAccessRegionEntity> matching = regionEntities(unitId).stream().filter(region -> sheetId.equals(region.getSheetId())).toList();
        matching.forEach(store::deleteAccessRegion);
        return !matching.isEmpty();
    }

    private boolean duplicateSheet(String unitId, String sourceSheetId, String targetSheetId, JsonNode snapshotAfter) {
        if (sourceSheetId == null || sourceSheetId.isBlank() || targetSheetId == null || targetSheetId.isBlank()) {
            throw ServiceException.validation("Duplicated sheet identities are invalid");
        }
        boolean targetExists = false;
        for (JsonNode sheet : snapshotAfter.path("sheets")) if (targetSheetId.equals(sheet.path("id").asText())) targetExists = true;
        if (!targetExists) throw ServiceException.validation("Duplicated sheet is missing from the structural result");
        List<RangeAccessRegion> source = loadRegions(unitId).stream().filter(region -> sourceSheetId.equals(region.sheetId())).toList();
        Instant now = Instant.now();
        for (RangeAccessRegion region : source) {
            String copiedId = UUID.randomUUID().toString();
            RangeRef copiedRange = new RangeRef(targetSheetId, region.range().startRow(), region.range().endRow(),
                    region.range().startColumn(), region.range().endColumn());
            RangeAccessRegionEntity copy = new RangeAccessRegionEntity(copiedId, unitId, targetSheetId, copiedRange,
                    region.defaultAccess(), region.createdBy(), now);
            store.saveAccessRegion(copy, grantEntities(copiedId, region.grants()));
        }
        return !source.isEmpty();
    }

    private void requireOwnerAndLockWorkbook(String unitId, String subject) {
        store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
        access.require(unitId, subject, WorkbookRole.OWNER);
    }

    private void validateRequest(String unitId, RangeAccessRegionRequest request) {
        if (request == null || request.sheetId() == null || request.sheetId().isBlank()
                || request.range() == null || request.defaultAccess() == null) {
            throw ServiceException.validation("Range access region requires a sheet, range, and default access");
        }
        RangeRef range = request.range();
        if (!request.sheetId().equals(range.sheetId()) || range.endRow() > MAX_ROW_INDEX || range.endColumn() > MAX_COLUMN_INDEX) {
            throw ServiceException.validation("Range access bounds or sheet identity are invalid");
        }
        WorkbookRow workbook = store.find(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
        try {
            JsonNode root = mapper.readTree(workbook.snapshotJson());
            boolean sheetExists = false;
            for (JsonNode sheet : root.path("sheets")) {
                if (request.sheetId().equals(sheet.path("id").asText())) { sheetExists = true; break; }
            }
            if (!sheetExists) throw ServiceException.validation("Range access sheet does not exist in the canonical workbook");
        } catch (ServiceException error) {
            throw error;
        } catch (Exception error) {
            throw new ServiceException("STORAGE_CORRUPT", 500, "Workbook snapshot cannot be read while validating range access", error);
        }
        Set<String> principals = new HashSet<>();
        for (RangeAccessGrant grant : request.grants()) {
            if (grant == null || grant.principal() == null || grant.principal().kind() == null || grant.access() == null) {
                throw new ServiceException("ACCESS_PRINCIPAL_INVALID", 400, "Range access grant is incomplete");
            }
            AccessPrincipal principal = grant.principal();
            String principalId = principal.id();
            if (principal.kind() == AccessPrincipalKind.EVERYONE) {
                if (principalId != null && !principalId.isBlank()) {
                    throw new ServiceException("ACCESS_PRINCIPAL_INVALID", 400, "Everyone principal cannot have an id");
                }
                principalId = "";
            } else if (principalId == null || principalId.isBlank() || !principalId.equals(principalId.trim()) || principalId.length() > 500) {
                throw new ServiceException("ACCESS_PRINCIPAL_INVALID", 400, "Subject and group principals require a valid id");
            }
            if (!principals.add(principal.kind().wireValue() + "\u0000" + principalId)) {
                throw new ServiceException("ACCESS_PRINCIPAL_INVALID", 400, "A principal may have only one grant in a range");
            }
        }
    }

    private void requireNoOverlap(String unitId, RangeRef candidate, String excludedId) {
        for (RangeAccessRegion existing : loadRegions(unitId)) {
            if (!existing.id().equals(excludedId) && RangeAccessIndex.intersects(candidate, existing.range())) {
                throw new ServiceException("ACCESS_REGION_OVERLAP", 409, "Range access regions on one sheet cannot overlap");
            }
        }
    }

    private List<RangeAccessRegionEntity> regionEntities(String unitId) {
        return store.listAccessRegions(unitId);
    }

    private List<RangeAccessRegion> loadRegions(String unitId) {
        List<RangeAccessRegionEntity> entities = regionEntities(unitId);
        return mapRegions(entities, store.listAccessGrants(entities.stream().map(RangeAccessRegionEntity::getId).toList()));
    }

    private AccessIndexSnapshot loadIndex(String unitId, long revision) {
        List<RangeAccessRegionEntity> entities = regionEntities(unitId);
        List<RangeAccessRegion> regions = mapRegions(entities,
                store.listAccessGrants(entities.stream().map(RangeAccessRegionEntity::getId).toList()));
        return new AccessIndexSnapshot(revision, regions, new RangeAccessIndex(regions));
    }

    private List<RangeAccessRegion> mapRegions(List<RangeAccessRegionEntity> regions, List<RangeAccessGrantEntity> grants) {
        Map<String, List<RangeAccessGrant>> grantsByRegion = new HashMap<>();
        for (RangeAccessGrantEntity grant : grants) {
            RangeAccessGrantEntity.Id id = grant.getId();
            AccessPrincipal principal = new AccessPrincipal(id.getPrincipalKind(),
                    id.getPrincipalKind() == AccessPrincipalKind.EVERYONE ? null : id.getPrincipalId());
            grantsByRegion.computeIfAbsent(id.getRegionId(), ignored -> new ArrayList<>()).add(new RangeAccessGrant(principal, grant.getAccess()));
        }
        return regions.stream().map(region -> new RangeAccessRegion(region.getId(), region.getUnitId(), region.getSheetId(),
                region.getRange(), region.getDefaultAccess(), grantsByRegion.getOrDefault(region.getId(), List.of()),
                region.getCreatedBy(), region.getCreatedAt(), region.getUpdatedAt())).toList();
    }

    private List<RangeAccessGrantEntity> grantEntities(String regionId, List<RangeAccessGrant> grants) {
        return grants.stream().map(grant -> new RangeAccessGrantEntity(regionId, grant.principal().kind(),
                grant.principal().kind() == AccessPrincipalKind.EVERYONE ? "" : grant.principal().id(), grant.access())).toList();
    }

    private RangeAccessRegion toRegion(RangeAccessRegionEntity entity, List<RangeAccessGrant> grants) {
        return new RangeAccessRegion(entity.getId(), entity.getUnitId(), entity.getSheetId(), entity.getRange(),
                entity.getDefaultAccess(), grants, entity.getCreatedBy(), entity.getCreatedAt(), entity.getUpdatedAt());
    }

    private void evict(String unitId) {
        indexes.keySet().removeIf(key -> key.unitId().equals(unitId));
    }

    private record IndexKey(String unitId, long revision) { }
    private record AccessIndexSnapshot(long revision, List<RangeAccessRegion> regions, RangeAccessIndex index) { }
}
