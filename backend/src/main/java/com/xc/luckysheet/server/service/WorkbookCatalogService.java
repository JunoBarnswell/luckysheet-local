package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.CopyWorkbookRequest;
import com.xc.luckysheet.server.contract.CursorPage;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.GeneratedWorkbookContract;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.UpdateWorkbookRequest;
import com.xc.luckysheet.server.contract.UserStateRequest;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookArtifactResponse;
import com.xc.luckysheet.server.contract.WorkbookImportResponse;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.contract.WorkbookOpenResponse;
import com.xc.luckysheet.server.contract.WorkbookSource;
import com.xc.luckysheet.server.contract.WorkbookSummary;
import com.xc.luckysheet.server.contract.WorkbookSyncStatus;
import com.xc.luckysheet.server.contract.WorkbookUserState;
import com.xc.luckysheet.server.config.KernelHostProperties;
import com.xc.luckysheet.server.persistence.AuditEntityRepository;
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
import java.nio.file.Files;
import java.nio.file.Path;
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
    public static final long MAX_NATIVE_DOCUMENT_BYTES = 1024L * 1024L * 1024L;
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
    private final OperationEntityRepository operationEntities;
    private final OutboxEntityRepository outbox;
    private final AuditEntityRepository audits;
    private final ShareEntityRepository shares;
    private final DataBlockEntityRepository blocks;
    private final ObjectMapper mapper;
    private final KernelHostClient kernel;
    private final KernelPersistenceService kernelPersistence;
    private final KernelHostProperties kernelProperties;

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
            OperationEntityRepository operationEntities,
            OutboxEntityRepository outbox,
            AuditEntityRepository audits,
            ShareEntityRepository shares,
            DataBlockEntityRepository blocks,
            ObjectMapper mapper, KernelHostClient kernel, KernelPersistenceService kernelPersistence, KernelHostProperties kernelProperties
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
        this.operationEntities = operationEntities;
        this.outbox = outbox;
        this.audits = audits;
        this.shares = shares;
        this.blocks = blocks;
        this.mapper = mapper;
        this.kernel = kernel; this.kernelPersistence = kernelPersistence; this.kernelProperties = kernelProperties;
    }

    @Transactional
    public WorkbookOpenResponse create(CreateWorkbookRequest request, String actor) {
        if (workbooks.existsById(request.unitId())) throw ServiceException.conflict("Workbook already exists");
        ObjectNode params = mapper.createObjectNode().put("unitId", request.unitId()).put("name", request.name().trim());
        if (request.sheets() != null) params.set("sheets", request.sheets());
        JsonNode manifest;
        synchronized (kernel) {
            invalidateOnRollback();
            manifest = kernel.call("create", params);
        }
        WorkbookEntity entity = createNativeEntity(request.unitId(), manifest, request.spaceId(), request.folderId(), WorkbookSource.NATIVE, actor);
        ObjectNode result = mapper.createObjectNode(); result.set("manifest", manifest); result.putArray("pages");
        kernelPersistence.publish(result, request.unitId(), 0);
        if (request.initialMutations().isEmpty()) return openResponse(entity, manifest);
        OperationEnvelope initialOperation = new OperationEnvelope(
                OperationEnvelope.SCHEMA,
                "workbook-create-" + UUID.randomUUID(),
                request.unitId(),
                1,
                0,
                request.initialMutations(),
                Instant.now()
        );
        var committed = operations.commit(request.unitId(), initialOperation, actor);
        WorkbookEntity committedEntity = workbooks.findById(request.unitId())
                .orElseThrow(() -> ServiceException.notFound("Workbook creation was not persisted"));
        return openResponse(committedEntity, committed.changeSet().path("manifest"));
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
        WorkbookEntity source = lockActiveOrTrashed(unitId);
        if (source.getLifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash");
        requireRole(unitId, actor, WorkbookAclRole.VIEWER);
        String spaceId = request == null || request.spaceId() == null ? source.getSpaceId() : blankToNull(request.spaceId());
        String folderId = request == null || request.folderId() == null ? source.getFolderId() : blankToNull(request.folderId());
        String name = request == null || request.name() == null || request.name().isBlank() ? source.getName() + " - 副本" : request.name().trim();
        String targetId = UUID.randomUUID().toString();
        ObjectNode params = mapper.createObjectNode().put("sourceUnitId", unitId).put("sourceRevision", source.getRevision())
                .put("targetUnitId", targetId).put("name", name);
        JsonNode result;
        synchronized (kernel) {
            invalidateOnRollback();
            kernelPersistence.reopen(unitId, source.getRevision(), kernel);
            result = kernel.call("copy", params);
        }
        JsonNode manifest = result.path("manifest");
        WorkbookEntity copied = createNativeEntity(targetId, manifest, spaceId, folderId, source.getSource(), actor);
        kernelPersistence.copyPages(unitId, source.getRevision(), targetId, manifest);
        WorkbookSourceArtifactEntity sourceArtifact = artifacts.findById(unitId).orElse(null);
        if (sourceArtifact != null) copyNativeArtifact(targetId, sourceArtifact);
        return summaryForActor(copied, actor);
    }

    private void copyNativeArtifact(String targetId, WorkbookSourceArtifactEntity source) {
        Path directory = null;
        try {
            directory = Files.createTempDirectory(taskRoot(), "native-copy-");
            Path pagesDirectory = Files.createDirectory(directory.resolve("pages"));
            Path output = directory.resolve("copy.artifact");
            // Copy is an explicit identity boundary: bind the preserved package to
            // revision zero, then export every copied canonical page into that package.
            ObjectNode params = mapper.createObjectNode().put("unitId", targetId).put("revision", 0)
                    .put("fileHandle", output.toString()).put("format", source.getFormat())
                    .put("sourceFileHandle", verifiedArtifactPath(source).toString())
                    .put("sourceChecksum", source.getChecksum()).put("sourceRevision", 0)
                    .put("pagesDirectory", pagesDirectory.toString());
            JsonNode result;
            synchronized (kernel) {
                kernelPersistence.reopen(targetId, 0, kernel);
                kernelPersistence.stagePages(targetId, 0, pagesDirectory);
                result = kernel.call("document.export", params);
            }
            storeArtifact(targetId, 0, source.getFileName(), output, result.path("artifact"));
        } catch (IOException error) {
            throw ServiceException.unavailable("Native copy I/O failed: " + error.getMessage());
        } finally {
            deleteTaskDirectory(directory);
        }
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
        kernelPersistence.purge(unitId);
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

    /** Artifacts are exclusively generated by native from a committed revision. */
    @Transactional
    public WorkbookArtifactResponse exportArtifact(String unitId, long revision, String fileName, String format, String actor) {
        WorkbookEntity workbook = lockActiveOrTrashed(unitId);
        requireRole(unitId, actor, WorkbookAclRole.EDITOR);
        if (workbook.getLifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash");
        if (workbook.getRevision() != revision) throw ServiceException.conflict("Artifact revision must match current workbook revision");
        Path root = taskRoot();
        Path output = null;
        Path exportDirectory = null;
        Path pagesDirectory = null;
        try {
            exportDirectory = Files.createTempDirectory(root, "native-export-");
            output = exportDirectory.resolve("export.artifact");
            WorkbookSourceArtifactEntity previous = artifacts.findById(unitId).orElse(null);
            ObjectNode params = mapper.createObjectNode().put("unitId", unitId).put("revision", revision)
                    .put("fileHandle", output.toString()).put("format", format);
            if (previous != null) {
                params.put("sourceFileHandle", verifiedArtifactPath(previous).toString());
                params.put("sourceChecksum", previous.getChecksum());
                params.put("sourceRevision", previous.getWorkbookRevision());
            }
            JsonNode result;
            synchronized (kernel) {
                kernelPersistence.reopen(unitId, revision, kernel);
                pagesDirectory = Files.createTempDirectory(root, "native-export-pages-");
                kernelPersistence.stagePages(unitId, revision, pagesDirectory);
                params.put("pagesDirectory", pagesDirectory.toString());
                result = kernel.call("document.export", params);
            }
            JsonNode metadata = artifactMetadata(result);
            if (metadata.path("revision").asLong(-1) != revision) throw new KernelHostException("REVISION_CONFLICT", "Native artifact revision mismatch", unitId, "regenerate-artifact");
            WorkbookSourceArtifactEntity saved = storeArtifact(unitId, revision, safeFileName(fileName), output, metadata);
            output = null;
            return artifactResponse(saved);
        } catch (IOException error) { throw ServiceException.unavailable("Native export I/O failed: " + error.getMessage()); }
        finally { deleteTemporary(output); deleteTaskDirectory(pagesDirectory); deleteTaskDirectory(exportDirectory); }
    }

    public WorkbookSourceArtifactEntity getArtifact(String unitId, String actor) {
        requireRole(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookEntity workbook = requireActiveOrTrashed(unitId);
        WorkbookSourceArtifactEntity artifact = artifacts.findById(unitId).orElseThrow(() -> ServiceException.notFound("Workbook native artifact not found"));
        if (artifact.getWorkbookRevision() != workbook.getRevision()) throw ServiceException.conflict("Artifact is stale; export the current revision first");
        verifiedArtifactPath(artifact);
        return artifact;
    }

    @Transactional
    public WorkbookImportResponse importNativePath(Path input, String name, String spaceId, String folderId, String actor,
                                                    java.util.function.BooleanSupplier cancelled) {
        Path root = taskRoot();
        Path directory = null;
        Path retained = null;
        try {
            Path file = input.toRealPath();
            if (!file.startsWith(root) || !Files.isRegularFile(file)) throw ServiceException.forbidden("Import file is outside native task storage");
            long length = Files.size(file);
            if (length < 1 || length > MAX_NATIVE_DOCUMENT_BYTES) throw ServiceException.validation("Native import must be between 1 byte and 1 GiB");
            String unitId = UUID.randomUUID().toString();
            String resolvedName = name == null || name.isBlank() ? file.getFileName().toString() : name.trim();
            directory = Files.createTempDirectory(root, "native-import-pages-");
            ObjectNode params = mapper.createObjectNode().put("unitId", unitId).put("name", resolvedName)
                    .put("fileHandle", file.toString()).put("outputDirectory", directory.toString());
            JsonNode result;
            synchronized (kernel) { invalidateOnRollback(); result = kernel.call("document.import", params); }
            JsonNode manifest = result.path("manifest");
            WorkbookEntity entity = createNativeEntity(unitId, manifest, spaceId, folderId, WorkbookSource.DOCUMENT_IMPORT, actor);
            kernelPersistence.publishImported(result, unitId, 0, root);
            JsonNode metadata = artifactMetadata(result);
            retained = Files.createTempFile(root, "source-artifact-", ".artifact");
            Files.copy(file, retained, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
            WorkbookSourceArtifactEntity artifact = storeArtifact(unitId, 0, resolvedName, retained, metadata);
            org.springframework.transaction.support.TransactionSynchronizationManager.registerSynchronization(new org.springframework.transaction.support.TransactionSynchronization() {
                @Override public void beforeCommit(boolean readOnly) {
                    if (cancelled.getAsBoolean()) throw new ServiceException("TASK_CANCELLED", 409, "Import task was cancelled before database commit");
                }
            });
            return new WorkbookImportResponse(unitId, 0, artifact.getChecksum(), summaryForActor(entity, actor), manifest, artifactResponse(artifact));
        } catch (IOException error) { throw new KernelHostException("DOCUMENT_IMPORT_FAILED", error.getMessage(), input.toString(), "retry-import-after-repairing-task-storage"); }
        finally { deleteTemporary(retained); deleteTaskDirectory(directory); }
    }

    private WorkbookEntity createNativeEntity(String unitId, JsonNode manifest, String spaceId, String folderId, WorkbookSource source, String actor) {
        if (!manifest.isObject() || manifest.path("version").asInt(-1) != 11 || manifest.path("revision").asLong(-1) != 0 || !unitId.equals(manifest.path("unitId").asText()))
            throw new KernelHostException("MANIFEST_INVALID", "Native creation must return the requested v11 resource at revision zero", unitId, "reject-creation");
        if (workbooks.existsById(unitId)) throw ServiceException.conflict("Workbook already exists");
        WorkspaceSpaceEntity space = spaceId == null || spaceId.isBlank() ? workspace.ensurePersonalSpace(actor) : workspace.require(spaceId, actor, WorkbookAclRole.EDITOR);
        String normalizedFolder = blankToNull(folderId);
        workspace.requireFolder(space.getSpaceId(), normalizedFolder, actor, WorkbookAclRole.EDITOR);
        Instant now = Instant.now();
        WorkbookEntity entity = new WorkbookEntity(unitId, manifest.path("name").asText(), 0, now, now, actor, space.getSpaceId(), normalizedFolder,
                com.xc.luckysheet.server.contract.WorkbookStorageLocation.REMOTE, source, WorkbookLifecycle.ACTIVE, null);
        workbooks.save(entity);
        acl.save(new WorkbookAclEntity(unitId, actor, WorkbookAclRole.OWNER, now, now));
        return entity;
    }

    public Path verifiedArtifactPath(WorkbookSourceArtifactEntity artifact) {
        try {
            Path path = Path.of(artifact.getStoragePath()).toRealPath();
            if (!path.startsWith(taskRoot()) || !Files.isRegularFile(path) || Files.size(path) != artifact.getByteLength() || !checksumFile(path).equals(artifact.getChecksum()))
                throw ServiceException.conflict("Native artifact storage does not match its durable metadata");
            return path;
        } catch (IOException error) { throw ServiceException.unavailable("Native artifact is unavailable: " + error.getMessage()); }
    }

    private WorkbookSourceArtifactEntity storeArtifact(String unitId, long revision, String fileName, Path file, JsonNode metadata) throws IOException {
        String digest = checksumFile(file);
        if (metadata.path("revision").asLong(-1) != revision || !digest.equals(metadata.path("checksum").asText())
                || Files.size(file) != metadata.path("byteLength").asLong(-1) || metadata.path("format").asText().isBlank())
            throw new KernelHostException("ARTIFACT_INVALID", "Native artifact identity, size or checksum does not match bytes", unitId, "reject-artifact-publication");
        Path durableDirectory = taskRoot().resolve("committed-artifacts");
        Files.createDirectories(durableDirectory);
        Path durable = durableDirectory.resolve(UUID.randomUUID() + ".artifact");
        Files.move(file, durable, java.nio.file.StandardCopyOption.ATOMIC_MOVE);
        org.springframework.transaction.support.TransactionSynchronizationManager.registerSynchronization(new org.springframework.transaction.support.TransactionSynchronization() {
            @Override public void afterCompletion(int status) { if (status != STATUS_COMMITTED) deleteTemporary(durable); }
        });
        Instant now = Instant.now();
        WorkbookSourceArtifactEntity entity = artifacts.findById(unitId).orElseGet(() -> new WorkbookSourceArtifactEntity(unitId, fileName,
                "application/octet-stream", digest, revision, metadata.path("byteLength").asLong(), durable.toString(), writeJson(metadata), now, now));
        entity.update(fileName, "application/octet-stream", digest, revision, metadata.path("byteLength").asLong(), durable.toString(), writeJson(metadata), now);
        artifacts.save(entity);
        return entity;
    }

    private Path taskRoot() {
        try {
            if (kernelProperties.taskDirectory() == null || kernelProperties.taskDirectory().isBlank()) throw ServiceException.unavailable("KERNEL_HOST_TASK_DIRECTORY is required");
            Path root = Path.of(kernelProperties.taskDirectory()).toAbsolutePath().normalize();
            Files.createDirectories(root);
            return root.toRealPath();
        } catch (IOException error) { throw ServiceException.unavailable("Native task storage is unavailable"); }
    }
    private String checksumFile(Path path) throws IOException {
        try {
            MessageDigest digest = MessageDigest.getInstance("SHA-256");
            try (var input = Files.newInputStream(path)) { byte[] buffer = new byte[65536]; int read; while ((read = input.read(buffer)) != -1) digest.update(buffer, 0, read); }
            return HexFormat.of().formatHex(digest.digest());
        } catch (java.security.NoSuchAlgorithmException error) { throw new IllegalStateException(error); }
    }
    private void deleteTaskDirectory(Path directory) {
        if (directory == null) return;
        Path normalized = directory.toAbsolutePath().normalize();
        if (!normalized.startsWith(taskRoot()) || normalized.equals(taskRoot())) throw new IllegalStateException("Task cleanup outside task storage");
        try (var entries = Files.walk(normalized)) {
            for (Path path : entries.sorted(java.util.Comparator.reverseOrder()).toList()) Files.deleteIfExists(path);
        } catch (IOException error) { org.slf4j.LoggerFactory.getLogger(getClass()).error("Native task directory cleanup failed: {}", normalized, error); }
    }
    private void deleteTemporary(Path path) {
        if (path == null) return;
        try { Files.deleteIfExists(path); } catch (IOException error) { org.slf4j.LoggerFactory.getLogger(getClass()).error("Native temporary file cleanup failed: {}", path, error); }
    }
    private void invalidateOnRollback() {
        if (!org.springframework.transaction.support.TransactionSynchronizationManager.isSynchronizationActive()) throw new IllegalStateException("Native mutation requires transaction");
        org.springframework.transaction.support.TransactionSynchronizationManager.registerSynchronization(new org.springframework.transaction.support.TransactionSynchronization() {
            @Override public void afterCompletion(int status) { if (status != STATUS_COMMITTED) kernel.abortTransaction(); }
        });
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

    private WorkbookOpenResponse openResponse(WorkbookEntity entity, JsonNode manifest) {
        String json = writeJson(manifest);
        return new WorkbookOpenResponse(entity.getUnitId(), entity.getRevision(), manifest, checksum(json.getBytes(StandardCharsets.UTF_8)));
    }

    private WorkbookUserState userState(WorkbookUserStateEntity state) {
        return new WorkbookUserState(state.getId().getUnitId(), state.isFavorite(), state.getLastOpenedAt(), state.isAutoSave(),
                state.isAutoSync(), state.getDefaultCreateLocation(), state.getImportCompatibilityLevel(), state.getLanguage(),
                state.isOfflineCache(), state.getTheme(), state.getUpdatedAt());
    }

    private WorkbookArtifactResponse artifactResponse(WorkbookSourceArtifactEntity artifact) {
        JsonNode nativeMetadata;
        try { nativeMetadata = mapper.readTree(artifact.getNativeMetadataJson()); }
        catch (Exception error) { throw new KernelHostException("ARTIFACT_INVALID", "Stored native artifact metadata is invalid", artifact.getUnitId(), "reimport-or-regenerate-artifact"); }
        return new WorkbookArtifactResponse(artifact.getUnitId(), artifact.getFileName(), artifact.getMimeType(), artifact.getChecksum(), artifact.getWorkbookRevision(),
                artifact.getByteLength(), nativeMetadata, artifact.getCreatedAt(), artifact.getUpdatedAt());
    }

    private JsonNode artifactMetadata(JsonNode nativeResult) {
        if (!nativeResult.path("artifact").isObject() || !nativeResult.path("metadata").isObject())
            throw new KernelHostException("ARTIFACT_INVALID", "Native result omitted artifact or document metadata", null, "deploy-matching-kernel-host");
        ObjectNode metadata = ((ObjectNode) nativeResult.path("artifact")).deepCopy();
        metadata.set("documentMetadata", nativeResult.path("metadata").deepCopy());
        return metadata;
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

    private String safeFileName(String value) {
        if (value == null || value.isBlank()) throw ServiceException.validation("Native document file name is required");
        try {
            return java.net.URLDecoder.decode(value, java.nio.charset.StandardCharsets.UTF_8).replaceAll("[\\r\\n]", "_");
        } catch (IllegalArgumentException ignored) {
            return value.replaceAll("[\\r\\n]", "_");
        }
    }
    private String safeMimeType(String value) { return value == null || value.isBlank() ? "application/octet-stream" : value; }
    private String writeJson(Object value) { try { return mapper.writeValueAsString(value); } catch (Exception error) { throw new IllegalStateException("Unable to serialize workbook snapshot", error); } }
    private String checksum(byte[] content) { try { return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content)); } catch (Exception error) { throw new IllegalStateException("SHA-256 is unavailable", error); } }
    private boolean nonBlank(String value) { return value != null && !value.isBlank(); }
    private static String blankToNull(String value) { return value == null || value.isBlank() ? null : value; }
}
