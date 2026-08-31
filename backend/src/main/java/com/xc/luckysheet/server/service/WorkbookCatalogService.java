package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.CopyWorkbookRequest;
import com.xc.luckysheet.server.contract.CursorPage;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.GeneratedWorkbookContract;
import com.xc.luckysheet.server.contract.UpdateWorkbookRequest;
import com.xc.luckysheet.server.contract.UserStateRequest;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookArtifactResponse;
import com.xc.luckysheet.server.contract.WorkbookImportResponse;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.contract.WorkbookSnapshotResponse;
import com.xc.luckysheet.server.contract.WorkbookSnapshotValidator;
import com.xc.luckysheet.server.contract.WorkbookSource;
import com.xc.luckysheet.server.contract.WorkbookSummary;
import com.xc.luckysheet.server.contract.WorkbookSyncStatus;
import com.xc.luckysheet.server.contract.WorkbookUserState;
import com.xc.luckysheet.server.persistence.AuditEntityRepository;
import com.xc.luckysheet.server.persistence.CheckpointEntityRepository;
import com.xc.luckysheet.server.persistence.DataBlockEntityRepository;
import com.xc.luckysheet.server.persistence.OperationEntityRepository;
import com.xc.luckysheet.server.persistence.OutboxEntityRepository;
import com.xc.luckysheet.server.persistence.ShareEntityRepository;
import com.xc.luckysheet.server.persistence.SpaceMemberEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookAclEntity;
import com.xc.luckysheet.server.persistence.WorkbookAclEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookEntity;
import com.xc.luckysheet.server.persistence.WorkbookEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookSourceArtifactEntity;
import com.xc.luckysheet.server.persistence.WorkbookSourceArtifactEntityRepository;
import com.xc.luckysheet.server.persistence.WorkbookUserStateEntity;
import com.xc.luckysheet.server.persistence.WorkbookUserStateEntityRepository;
import com.xc.luckysheet.server.persistence.WorkspaceFolderEntity;
import com.xc.luckysheet.server.persistence.WorkspaceFolderEntityRepository;
import com.xc.luckysheet.server.persistence.WorkspaceSpaceEntity;
import com.xc.luckysheet.server.persistence.WorkspaceSpaceEntityRepository;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.data.domain.PageRequest;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Instant;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import java.util.HexFormat;

/** The sole backend catalog boundary for workbook resources and their artifacts. */
@Service
public class WorkbookCatalogService {
    public static final long MAX_NATIVE_DOCUMENT_BYTES = 50L * 1024L * 1024L;
    /** The server accepts one canonical exchange contract at a time. */
    public static final int NATIVE_DOCUMENT_CODEC_REVISION = 1;

    private final WorkbookEntityRepository workbooks;
    private final WorkbookAclEntityRepository acl;
    private final WorkbookUserStateEntityRepository userStates;
    private final WorkbookSourceArtifactEntityRepository artifacts;
    private final WorkspaceSpaceEntityRepository spaces;
    private final WorkspaceFolderEntityRepository folders;
    private final SpaceMemberEntityRepository members;
    private final WorkspaceService workspace;
    private final WorkbookAuthorizationService authorization;
    private final WorkbookOperationService operations;
    private final CheckpointEntityRepository checkpoints;
    private final OperationEntityRepository operationEntities;
    private final OutboxEntityRepository outbox;
    private final AuditEntityRepository audits;
    private final ShareEntityRepository shares;
    private final DataBlockEntityRepository blocks;
    private final ObjectMapper mapper;

    public WorkbookCatalogService(
            WorkbookEntityRepository workbooks,
            WorkbookAclEntityRepository acl,
            WorkbookUserStateEntityRepository userStates,
            WorkbookSourceArtifactEntityRepository artifacts,
            WorkspaceSpaceEntityRepository spaces,
            WorkspaceFolderEntityRepository folders,
            SpaceMemberEntityRepository members,
            WorkspaceService workspace,
            WorkbookAuthorizationService authorization,
            WorkbookOperationService operations,
            CheckpointEntityRepository checkpoints,
            OperationEntityRepository operationEntities,
            OutboxEntityRepository outbox,
            AuditEntityRepository audits,
            ShareEntityRepository shares,
            DataBlockEntityRepository blocks,
            ObjectMapper mapper
    ) {
        this.workbooks = workbooks;
        this.acl = acl;
        this.userStates = userStates;
        this.artifacts = artifacts;
        this.spaces = spaces;
        this.folders = folders;
        this.members = members;
        this.workspace = workspace;
        this.authorization = authorization;
        this.operations = operations;
        this.checkpoints = checkpoints;
        this.operationEntities = operationEntities;
        this.outbox = outbox;
        this.audits = audits;
        this.shares = shares;
        this.blocks = blocks;
        this.mapper = mapper;
    }

    @Transactional
    public WorkbookSnapshotResponse create(CreateWorkbookRequest request, String actor) {
        WorkbookEntity entity = createEntity(request, actor);
        return snapshotResponse(entity, request.snapshot());
    }

    public CursorPage<WorkbookSummary> list(String actor, String view, String spaceId, String folderId, String query, int page, int limit) {
        String normalizedView = view == null || view.isBlank() ? "recent" : view.trim().toLowerCase();
        boolean trashOnly = "trash".equals(normalizedView);
        boolean includeTrash = trashOnly;
        boolean sharedOnly = "shared".equals(normalizedView);
        boolean ownedOnly = "owned".equals(normalizedView);
        String normalizedQuery = query == null || query.isBlank() ? null : query.trim();
        List<WorkbookEntity> rows = workbooks.findCatalogCandidates(actor, includeTrash, trashOnly, sharedOnly, ownedOnly,
                blankToNull(spaceId), blankToNull(folderId), normalizedQuery, PageRequest.of(page, limit));
        if (rows.isEmpty()) return new CursorPage<>(List.of(), null);

        List<String> unitIds = rows.stream().map(WorkbookEntity::getUnitId).toList();
        List<String> spaceIds = rows.stream().map(WorkbookEntity::getSpaceId).filter(this::nonBlank).distinct().toList();
        List<String> folderIds = rows.stream().map(WorkbookEntity::getFolderId).filter(this::nonBlank).distinct().toList();
        Map<String, WorkbookAclRole> directRoles = new HashMap<>();
        acl.findForSubjectAndUnits(actor, unitIds).forEach(item -> directRoles.put(item.getId().getUnitId(), item.getRole()));
        Map<String, WorkbookAclRole> spaceRoles = new HashMap<>();
        if (!spaceIds.isEmpty()) workspaceMembers(spaceIds, actor).forEach(item -> spaceRoles.put(item.getId().getSpaceId(), item.getRole()));
        Map<String, WorkspaceSpaceEntity> spaceMap = new HashMap<>();
        if (!spaceIds.isEmpty()) spaces.findAllById(spaceIds).forEach(item -> spaceMap.put(item.getSpaceId(), item));
        Map<String, WorkspaceFolderEntity> folderMap = new HashMap<>();
        if (!spaceIds.isEmpty()) folders.findBySpaceIdInOrderByName(spaceIds).forEach(item -> folderMap.put(item.getFolderId(), item));
        Map<String, WorkbookUserStateEntity> stateMap = new HashMap<>();
        userStates.findByIdUnitIdInAndIdSubject(unitIds, actor).forEach(item -> stateMap.put(item.getId().getUnitId(), item));
        Map<String, String> artifactNames = new HashMap<>();
        artifacts.findByUnitIdIn(unitIds).forEach(item -> artifactNames.put(item.getUnitId(), item.getFileName()));

        List<WorkbookSummary> items = rows.stream().map(row -> {
            WorkbookAclRole role = row.getOwnerSubject().equals(actor) ? WorkbookAclRole.OWNER : directRoles.get(row.getUnitId());
            role = max(role, row.getSpaceId() == null ? null : spaceRoles.get(row.getSpaceId()));
            if (role == null) role = WorkbookAclRole.VIEWER;
            WorkbookUserStateEntity state = stateMap.get(row.getUnitId());
            WorkspaceSpaceEntity space = row.getSpaceId() == null ? null : spaceMap.get(row.getSpaceId());
            WorkspaceFolderEntity folder = row.getFolderId() == null ? null : folderMap.get(row.getFolderId());
            return summary(row, role, state, space, folder, folderMap, artifactNames.get(row.getUnitId()));
        }).toList();
        return new CursorPage<>(items, items.size() == limit ? CursorPageRequest.next(page + 1) : null);
    }

    @Transactional
    public WorkbookSummary update(String unitId, UpdateWorkbookRequest request, String actor) {
        WorkbookEntity entity = lockActiveOrTrashed(unitId);
        WorkbookAclRole current = requireRole(unitId, actor, WorkbookAclRole.EDITOR);
        if (entity.getLifecycle() == WorkbookLifecycle.TRASHED) throw ServiceException.trashed("Workbook is in trash and cannot be moved");
        String targetSpaceId = request.spaceIdSpecified() ? request.spaceId() : entity.getSpaceId();
        String targetFolderId = request.folderIdSpecified() ? request.folderId() : entity.getFolderId();
        boolean crossSpace = !java.util.Objects.equals(targetSpaceId, entity.getSpaceId());
        if (crossSpace && !current.includes(WorkbookAclRole.OWNER)) {
            throw ServiceException.forbidden("Only the workbook owner can move it across spaces");
        }
        if (targetSpaceId == null) {
            targetSpaceId = workspace.ensurePersonalSpace(actor).getSpaceId();
        }
        workspace.requireFolder(targetSpaceId, targetFolderId, actor, WorkbookAclRole.EDITOR);
        workspace.require(targetSpaceId, actor, WorkbookAclRole.EDITOR);
        Instant now = Instant.now();
        entity.updateLocation(targetSpaceId, targetFolderId, now);
        workbooks.save(entity);
        if (crossSpace) acl.deleteNonOwner(unitId, entity.getOwnerSubject());
        return summaryForActor(entity, actor);
    }

    @Transactional
    public WorkbookSummary copy(String unitId, CopyWorkbookRequest request, String actor) {
        WorkbookEntity source = requireActiveOrTrashed(unitId);
        if (source.getLifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash and cannot be copied");
        requireRole(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookSnapshotResponse sourceSnapshot = operations.readSnapshot(unitId, actor);
        String targetSpaceId = request == null || request.spaceId() == null ? source.getSpaceId() : blankToNull(request.spaceId());
        if (targetSpaceId == null) targetSpaceId = workspace.ensurePersonalSpace(actor).getSpaceId();
        workspace.require(targetSpaceId, actor, WorkbookAclRole.EDITOR);
        String targetFolderId = request == null || request.folderId() == null ? source.getFolderId() : blankToNull(request.folderId());
        workspace.requireFolder(targetSpaceId, targetFolderId, actor, WorkbookAclRole.EDITOR);
        String name = request == null || request.name() == null || request.name().isBlank()
                ? source.getName() + " - 副本" : request.name().trim();
        String targetId = UUID.randomUUID().toString();
        JsonNode copiedSnapshot = normalizeCopiedSnapshot(sourceSnapshot.snapshot(), targetId, name);
        CreateWorkbookRequest create = new CreateWorkbookRequest(targetId, name, copiedSnapshot, targetSpaceId,
                targetFolderId, source.getSource());
        WorkbookEntity copied = createEntity(create, actor);
        WorkbookSourceArtifactEntity sourceArtifact = artifacts.findById(unitId).orElse(null);
        if (sourceArtifact != null) {
            Instant now = Instant.now();
            artifacts.save(new WorkbookSourceArtifactEntity(targetId, sourceArtifact.getFileName(), sourceArtifact.getMimeType(),
                    sourceArtifact.getChecksum(), sourceArtifact.getByteLength(), sourceArtifact.getContent().clone(),
                    sourceArtifact.getNativeMetadataJson(), now, now));
        }
        return summaryForActor(copied, actor);
    }

    @Transactional
    public WorkbookSummary moveToTrash(String unitId, String actor) {
        WorkbookEntity entity = lockActiveOrTrashed(unitId);
        requireRole(unitId, actor, WorkbookAclRole.OWNER);
        if (entity.getLifecycle() == WorkbookLifecycle.TRASHED) return summaryForActor(entity, actor);
        entity.moveToTrash(Instant.now());
        workbooks.save(entity);
        return summaryForActor(entity, actor);
    }

    @Transactional
    public WorkbookSummary restoreFromTrash(String unitId, String actor) {
        WorkbookEntity entity = lockActiveOrTrashed(unitId);
        requireRole(unitId, actor, WorkbookAclRole.OWNER);
        if (entity.getLifecycle() != WorkbookLifecycle.TRASHED) return summaryForActor(entity, actor);
        entity.restoreFromTrash(Instant.now());
        workbooks.save(entity);
        return summaryForActor(entity, actor);
    }

    @Transactional
    public void purge(String unitId, String actor) {
        WorkbookEntity entity = lockActiveOrTrashed(unitId);
        requireRole(unitId, actor, WorkbookAclRole.OWNER);
        if (entity.getLifecycle() != WorkbookLifecycle.TRASHED) throw ServiceException.conflict("Workbook must be in trash before purge");
        artifacts.deleteById(unitId);
        userStates.deleteByIdUnitId(unitId);
        blocks.deleteByIdUnitId(unitId);
        checkpoints.deleteByIdUnitId(unitId);
        operationEntities.deleteByUnitId(unitId);
        outbox.deleteByUnitId(unitId);
        audits.deleteByUnitId(unitId);
        shares.deleteByUnitId(unitId);
        acl.deleteAll(acl.findAllForWorkbook(unitId));
        workbooks.deleteById(unitId);
    }

    public WorkbookUserState getUserState(String unitId, String actor) {
        requireRole(unitId, actor, WorkbookAclRole.VIEWER);
        return userStates.findByIdUnitIdAndIdSubject(unitId, actor)
                .map(this::userState)
                .orElseGet(() -> new WorkbookUserState(unitId, false, null, true, true, "remote", "standard", null, true, "system", null));
    }

    @Transactional
    public WorkbookUserState putUserState(String unitId, UserStateRequest request, String actor) {
        requireRole(unitId, actor, WorkbookAclRole.VIEWER);
        Instant now = Instant.now();
        WorkbookUserStateEntity state = userStates.findByIdUnitIdAndIdSubject(unitId, actor)
                .orElseGet(() -> new WorkbookUserStateEntity(unitId, actor, false, null, now));
        state.update(request.favorite(), request.lastOpenedAt(), request.autoSave(), request.autoSync(),
                request.defaultCreateLocation(), request.importCompatibilityLevel(), request.language(),
                request.offlineCache(), request.theme(), now);
        userStates.save(state);
        return userState(state);
    }

    @Transactional
    public WorkbookArtifactResponse putArtifact(String unitId, String fileName, String mimeType, String checksum,
                                                byte[] content, String actor) {
        requireActive(unitId);
        requireRole(unitId, actor, WorkbookAclRole.EDITOR);
        validateArtifact(fileName, checksum, content);
        String actual = checksum(content);
        if (!actual.equalsIgnoreCase(checksum)) throw ServiceException.validation("Native document artifact checksum mismatch");
        Instant now = Instant.now();
        WorkbookSourceArtifactEntity entity = artifacts.findById(unitId).orElseGet(() ->
                new WorkbookSourceArtifactEntity(unitId, safeFileName(fileName), safeMimeType(mimeType), actual,
                        content.length, content.clone(), nativeArtifactMetadata(fileName), now, now));
        entity.update(safeFileName(fileName), safeMimeType(mimeType), actual, content.length, content.clone(), nativeArtifactMetadata(fileName), now);
        artifacts.save(entity);
        return artifactResponse(entity);
    }

    public WorkbookSourceArtifactEntity getArtifact(String unitId, String actor) {
        requireRole(unitId, actor, WorkbookAclRole.VIEWER);
        return artifacts.findById(unitId).orElseThrow(() -> ServiceException.notFound("Workbook native document artifact not found"));
    }

    @Transactional
    public WorkbookImportResponse importWorkbook(MultipartFile file, String name, String spaceId, String folderId,
                                                 String snapshotJson, String format, String nativeMetadataJson, String actor) {
        if (file == null || file.isEmpty()) throw ServiceException.validation("Native document file is required");
        if (file.getSize() > MAX_NATIVE_DOCUMENT_BYTES) throw ServiceException.validation("Native document exceeds 50 MiB");
        if (snapshotJson == null || snapshotJson.isBlank()) throw ServiceException.validation("Parsed workbook snapshot is required");
        JsonNode snapshot;
        try {
            snapshot = mapper.readTree(snapshotJson);
        } catch (IOException error) {
            throw ServiceException.validation("Parsed workbook snapshot is invalid");
        }
        if (snapshot == null || !snapshot.isObject()) throw ServiceException.validation("Parsed workbook snapshot must be an object");
        byte[] content;
        try {
            content = file.getBytes();
        } catch (IOException error) {
            throw ServiceException.unavailable("Unable to read native document artifact");
        }
        if (content.length == 0 || content.length > MAX_NATIVE_DOCUMENT_BYTES) {
            throw ServiceException.validation("Native document size is invalid");
        }
        String resolvedName = name == null || name.isBlank() ? file.getOriginalFilename() : name;
        if (resolvedName == null || resolvedName.isBlank()) resolvedName = "导入的工作簿";
        String unitId = snapshot.path("unitId").asText("").trim();
        if (unitId.isBlank() || unitId.length() > 200) throw ServiceException.validation("Parsed workbook snapshot must contain a valid unitId");
        WorkbookSnapshotValidator.requireCanonical(snapshot, unitId);
        JsonNode nativeMetadata;
        try {
            nativeMetadata = mapper.readTree(nativeMetadataJson);
        } catch (Exception error) {
            throw ServiceException.validation("Native package metadata is invalid");
        }
        String digest = validateNativeImportBinding(nativeMetadata, format, content, snapshot);
        ObjectNode artifactMetadata = ((ObjectNode) nativeMetadata).deepCopy();
        artifactMetadata.put("format", format.trim().toLowerCase(java.util.Locale.ROOT));
        WorkbookEntity entity = createEntity(new CreateWorkbookRequest(unitId, resolvedName, snapshot, spaceId, folderId,
                WorkbookSource.DOCUMENT_IMPORT), actor);
        Instant now = Instant.now();
        WorkbookSourceArtifactEntity artifact = new WorkbookSourceArtifactEntity(unitId,
                safeFileName(file.getOriginalFilename() == null ? resolvedName + ".ssjson" : file.getOriginalFilename()),
                safeMimeType(file.getContentType()), digest, content.length, content, writeJson(artifactMetadata), now, now);
        artifacts.save(artifact);
        return new WorkbookImportResponse(entity.getUnitId(), entity.getRevision(), artifact.getChecksum(),
                summaryForActor(entity, actor), snapshot.deepCopy(), artifactResponse(artifact));
    }

    /**
     * Establishes the trust boundary for browser-parsed native documents.
     * The server hashes the actual multipart bytes and the exact canonical
     * snapshot; the browser cannot substitute a hash, format, or codec
     * revision that was not proven against this request.
     */
    private String validateNativeImportBinding(JsonNode rawMetadata, String requestedFormat, byte[] content, JsonNode snapshot) {
        if (requestedFormat == null || requestedFormat.isBlank() || rawMetadata == null || !rawMetadata.isObject()) {
            throw ServiceException.validation("Native document metadata is invalid");
        }
        ObjectNode metadata = (ObjectNode) rawMetadata;
        Set<String> allowed = Set.of("schema", "format", "codecRevision", "checksum", "byteLength", "sourceSnapshotHash", "detectedFeatures", "compatibility");
        metadata.fieldNames().forEachRemaining(key -> {
            if (!allowed.contains(key)) throw ServiceException.validation("Native document metadata contains an unsupported field: " + key);
        });
        if (!"NativeDocumentMetadata".equals(metadata.path("schema").asText())) {
            throw ServiceException.validation("Native document metadata schema is invalid");
        }
        String format = requestedFormat.trim().toLowerCase(java.util.Locale.ROOT);
        if (!isSupportedNativeFormat(format) || !format.equals(metadata.path("format").asText("").trim().toLowerCase(java.util.Locale.ROOT))) {
            throw ServiceException.validation("Native document format binding is invalid");
        }
        JsonNode codecRevision = metadata.get("codecRevision");
        if (codecRevision == null || !codecRevision.isIntegralNumber() || codecRevision.intValue() != NATIVE_DOCUMENT_CODEC_REVISION) {
            throw ServiceException.validation("Native document codecRevision is unsupported");
        }
        String actualChecksum = checksum(content);
        if (!metadata.path("checksum").isTextual() || !actualChecksum.equalsIgnoreCase(metadata.path("checksum").asText())) {
            throw ServiceException.validation("Native document file checksum binding is invalid");
        }
        if (!metadata.path("byteLength").canConvertToLong() || metadata.path("byteLength").longValue() != content.length) {
            throw ServiceException.validation("Native document byteLength binding is invalid");
        }
        String expectedSnapshotHash = nativeSnapshotHash(snapshot);
        if (!metadata.path("sourceSnapshotHash").isTextual() || !expectedSnapshotHash.equalsIgnoreCase(metadata.path("sourceSnapshotHash").asText())) {
            throw ServiceException.validation("Native document snapshot hash binding is invalid");
        }
        if (metadata.has("detectedFeatures") && !metadata.get("detectedFeatures").isArray()) {
            throw ServiceException.validation("Native document detectedFeatures is invalid");
        }
        if (metadata.has("compatibility") && (!metadata.get("compatibility").isObject()
                || !"CompatibilityReport".equals(metadata.get("compatibility").path("schema").asText()))) {
            throw ServiceException.validation("Native document compatibility report is invalid");
        }
        return actualChecksum;
    }

    private boolean isSupportedNativeFormat(String format) {
        String[] parts = format.split("/", -1);
        if (parts.length != 2 || parts[0].isBlank() || parts[1].isBlank()
                || !format.matches("[a-z0-9-]+/[a-z0-9-]+")) return false;
        return switch (parts[0]) {
            case "ooxml" -> Set.of("xlsx", "xlsm", "xltx", "xltm", "xlam").contains(parts[1]);
            case "xlsb" -> parts[1].equals("xlsb");
            case "biff" -> Set.of("xls", "xlt", "xla", "biff5", "xlw").contains(parts[1]);
            case "xmlss" -> parts[1].equals("xml");
            case "text" -> Set.of("csv", "txt", "prn", "dif", "sylk").contains(parts[1]);
            case "ods", "sjs", "ssjson", "dbf" -> parts[0].equals(parts[1]);
            case "works" -> parts[1].equals("xlr");
            case "web" -> Set.of("html", "mht").contains(parts[1]);
            case "presentation" -> Set.of("pdf", "xps").contains(parts[1]);
            default -> false;
        };
    }

    /** Same compact FNV-1a snapshot identity as the native exchange codec. */
    private String nativeSnapshotHash(JsonNode value) {
        ObjectNode identity = ((ObjectNode) value).deepCopy();
        identity.put("unitId", "");
        JsonNode printDocuments = identity.get("printDocuments");
        if (printDocuments == null || printDocuments.isNull()) {
            identity.putArray("printDocuments");
        } else if (printDocuments.isArray()) {
            for (JsonNode raw : printDocuments) if (raw.isObject()) ((ObjectNode) raw).put("unitId", "");
        }
        String text = writeJson(identity);
        int hash = 0x811c9dc5;
        for (int index = 0; index < text.length(); index++) {
            hash ^= text.charAt(index);
            hash *= 0x01000193;
        }
        return "fnv1a-" + String.format("%08x", hash);
    }

    private WorkbookEntity createEntity(CreateWorkbookRequest request, String actor) {
        if (workbooks.existsById(request.unitId())) throw ServiceException.conflict("Workbook already exists");
        WorkbookSnapshotValidator.requireCanonical(request.snapshot(), request.unitId());
        if (!request.name().trim().equals(request.snapshot().path("name").asText().trim())) {
            throw ServiceException.validation("Workbook name must match the canonical snapshot name");
        }
        WorkspaceSpaceEntity space = request.spaceId() == null || request.spaceId().isBlank()
                ? workspace.ensurePersonalSpace(actor)
                : workspace.require(request.spaceId(), actor, WorkbookAclRole.EDITOR);
        String folderId = blankToNull(request.folderId());
        workspace.requireFolder(space.getSpaceId(), folderId, actor, WorkbookAclRole.EDITOR);
        Instant now = Instant.now();
        WorkbookEntity entity = new WorkbookEntity(request.unitId(), request.name().trim(), writeJson(request.snapshot()), 0, 0,
                now, now, actor, space.getSpaceId(), folderId,
                com.xc.luckysheet.server.contract.WorkbookStorageLocation.REMOTE,
                request.source(), WorkbookLifecycle.ACTIVE, null);
        workbooks.save(entity);
        acl.save(new WorkbookAclEntity(request.unitId(), actor, WorkbookAclRole.OWNER, now, now));
        checkpoints.save(new com.xc.luckysheet.server.persistence.CheckpointEntity(request.unitId(), 0, entity.getSnapshotJson(),
                checksum(entity.getSnapshotJson().getBytes(StandardCharsets.UTF_8)), now));
        return entity;
    }

    /**
     * A copied workbook is a new resource identity. All workbook-scoped
     * embedded documents must follow that identity before persistence; this
     * prevents the copied print state from being rejected or remaining bound
     * to the source workbook at the next load.
     */
    private JsonNode normalizeCopiedSnapshot(JsonNode source, String targetUnitId, String targetName) {
        if (source == null || !source.isObject()) throw ServiceException.validation("Workbook snapshot must be an object");
        ObjectNode copy = ((ObjectNode) source).deepCopy();
        copy.put("unitId", targetUnitId);
        copy.put("name", targetName);
        JsonNode printDocuments = copy.get("printDocuments");
        if (printDocuments != null && !printDocuments.isNull()) {
            if (!printDocuments.isArray()) throw ServiceException.validation("printDocuments must be an array");
            for (JsonNode raw : (ArrayNode) printDocuments) {
                if (!raw.isObject()) throw ServiceException.validation("Print document must be an object");
                ((ObjectNode) raw).put("unitId", targetUnitId);
            }
        }
        return copy;
    }

    private WorkbookSummary summaryForActor(WorkbookEntity entity, String actor) {
        Map<String, WorkspaceFolderEntity> folderMap = new HashMap<>();
        if (entity.getSpaceId() != null) folders.findBySpaceIdOrderByName(entity.getSpaceId()).forEach(item -> folderMap.put(item.getFolderId(), item));
        return summary(entity, requireRole(entity.getUnitId(), actor, WorkbookAclRole.VIEWER),
                userStates.findByIdUnitIdAndIdSubject(entity.getUnitId(), actor).orElse(null),
                entity.getSpaceId() == null ? null : spaces.findById(entity.getSpaceId()).orElse(null),
                entity.getFolderId() == null ? null : folderMap.get(entity.getFolderId()), folderMap,
                artifacts.findById(entity.getUnitId()).map(WorkbookSourceArtifactEntity::getFileName).orElse(null));
    }

    private WorkbookSummary summary(WorkbookEntity row, WorkbookAclRole role, WorkbookUserStateEntity state,
                                    WorkspaceSpaceEntity space, WorkspaceFolderEntity folder,
                                    Map<String, WorkspaceFolderEntity> folderMap, String sourceFileName) {
        List<String> path = locationPath(space, folder, folderMap);
        return new WorkbookSummary(row.getUnitId(), row.getName(), row.getRevision(), row.getUpdatedAt(), role,
                blankToNull(row.getOwnerSubject()), row.getSpaceId(), row.getFolderId(), path,
                space == null ? null : space.getName(), sourceFileName, row.getStorageLocation(),
                WorkbookSyncStatus.SYNCED, row.getLifecycle(), row.getSource(), state != null && state.isFavorite(),
                state == null ? null : state.getLastOpenedAt(), row.getDeletedAt());
    }

    private List<String> locationPath(WorkspaceSpaceEntity space, WorkspaceFolderEntity folder,
                                      Map<String, WorkspaceFolderEntity> folderMap) {
        if (space == null) return List.of();
        java.util.LinkedList<String> path = new java.util.LinkedList<>();
        path.addFirst(space.getName());
        java.util.Set<String> seen = new java.util.HashSet<>();
        WorkspaceFolderEntity current = folder;
        while (current != null && seen.add(current.getFolderId())) {
            path.add(1, current.getName());
            current = current.getParentId() == null ? null : folderMap.get(current.getParentId());
        }
        return List.copyOf(path);
    }

    private WorkbookSnapshotResponse snapshotResponse(WorkbookEntity entity, JsonNode snapshot) {
        String json = writeJson(snapshot);
        return new WorkbookSnapshotResponse(entity.getUnitId(), snapshot.deepCopy(), entity.getRevision(), checksum(json.getBytes(StandardCharsets.UTF_8)));
    }

    private WorkbookUserState userState(WorkbookUserStateEntity state) {
        return new WorkbookUserState(state.getId().getUnitId(), state.isFavorite(), state.getLastOpenedAt(), state.isAutoSave(),
                state.isAutoSync(), state.getDefaultCreateLocation(), state.getImportCompatibilityLevel(), state.getLanguage(),
                state.isOfflineCache(), state.getTheme(), state.getUpdatedAt());
    }

    private WorkbookArtifactResponse artifactResponse(WorkbookSourceArtifactEntity artifact) {
        return new WorkbookArtifactResponse(artifact.getUnitId(), artifact.getFileName(), artifact.getMimeType(), artifact.getChecksum(),
                artifact.getByteLength(), artifact.getCreatedAt(), artifact.getUpdatedAt());
    }

    private WorkbookEntity requireActiveOrTrashed(String unitId) {
        if (unitId == null || unitId.isBlank() || unitId.length() > 200) throw ServiceException.validation("unitId is invalid");
        return workbooks.findById(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
    }

    private WorkbookEntity lockActiveOrTrashed(String unitId) {
        if (unitId == null || unitId.isBlank() || unitId.length() > 200) throw ServiceException.validation("unitId is invalid");
        return workbooks.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
    }

    private void requireActive(String unitId) {
        WorkbookEntity entity = requireActiveOrTrashed(unitId);
        if (entity.getLifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash");
    }

    private WorkbookAclRole requireRole(String unitId, String actor, WorkbookAclRole required) {
        WorkbookAclRole role = authorization.role(unitId, actor).orElse(null);
        if (role == null) throw ServiceException.forbidden("Workbook access denied");
        if (!role.includes(required)) throw ServiceException.forbidden("Workbook role " + required + " is required");
        return role;
    }

    private List<com.xc.luckysheet.server.persistence.SpaceMemberEntity> workspaceMembers(Collection<String> spaceIds, String actor) {
        return members.findByIdSpaceIdInAndIdSubject(spaceIds, actor);
    }

    private WorkbookAclRole max(WorkbookAclRole left, WorkbookAclRole right) {
        if (left == null) return right;
        if (right == null) return left;
        return left.includes(right) ? left : right;
    }

    private void validateArtifact(String fileName, String checksum, byte[] content) {
        if (content == null || content.length == 0 || content.length > MAX_NATIVE_DOCUMENT_BYTES) throw ServiceException.validation("Native document artifact size is invalid");
        if (checksum == null || !checksum.matches("[A-Fa-f0-9]{64}")) throw ServiceException.validation("Native document artifact checksum is invalid");
        if (fileName != null && fileName.length() > GeneratedWorkbookContract.MAX_WORKBOOK_NAME_LENGTH) {
            throw ServiceException.validation("Native document file name is too long");
        }
    }

    private String safeFileName(String value) {
        if (value == null || value.isBlank()) throw ServiceException.validation("Native document file name is required");
        try {
            return java.net.URLDecoder.decode(value, java.nio.charset.StandardCharsets.UTF_8).replaceAll("[\\r\\n]", "_");
        } catch (IllegalArgumentException ignored) {
            return value.replaceAll("[\\r\\n]", "_");
        }
    }
    private String safeMimeType(String value) { return value == null || value.isBlank() ? "application/octet-stream" : value; }
    private String nativeArtifactMetadata(String fileName) {
        String lower = safeFileName(fileName).toLowerCase(java.util.Locale.ROOT);
        String family = lower.endsWith(".xlsb") ? "xlsb" : lower.endsWith(".xls") || lower.endsWith(".xlt") || lower.endsWith(".xla") || lower.endsWith(".xlw") ? "biff" : lower.endsWith(".xlsm") || lower.endsWith(".xltm") || lower.endsWith(".xltx") || lower.endsWith(".xlam") || lower.endsWith(".xlsx") ? "ooxml" : lower.endsWith(".ods") ? "ods" : lower.endsWith(".sjs") ? "sjs" : lower.endsWith(".ssjson") ? "ssjson" : lower.endsWith(".xml") ? "xmlss" : "text";
        String variant = lower.endsWith(".slk") ? "sylk" : lower.contains(".") ? lower.substring(lower.lastIndexOf('.') + 1) : "ssjson";
        return "{\"schema\":\"NativeDocumentMetadata\",\"format\":\"" + family + "/" + variant + "\",\"codecRevision\":1}";
    }
    private String writeJson(Object value) { try { return mapper.writeValueAsString(value); } catch (Exception error) { throw new IllegalStateException("Unable to serialize workbook snapshot", error); } }
    private String checksum(byte[] content) { try { return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content)); } catch (Exception error) { throw new IllegalStateException("SHA-256 is unavailable", error); } }
    private boolean nonBlank(String value) { return value != null && !value.isBlank(); }
    private static String blankToNull(String value) { return value == null || value.isBlank() ? null : value; }
}
