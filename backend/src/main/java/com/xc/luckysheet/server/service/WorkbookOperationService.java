package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.AclEntry;
import com.xc.luckysheet.server.contract.AuditRecord;
import com.xc.luckysheet.server.contract.CheckpointResponse;
import com.xc.luckysheet.server.contract.CursorPage;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.GeneratedWorkbookContract;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.OperationIntent;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.OperationOrigin;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.RestoreRequest;
import com.xc.luckysheet.server.contract.RevisionRecord;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookAccessProjection;
import com.xc.luckysheet.server.contract.WorkbookSnapshotResponse;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.contract.WorkbookSnapshotValidator;
import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.mutation.MutationDescriptorRegistry;
import com.xc.luckysheet.server.mutation.MutationPreparation;
import com.xc.luckysheet.server.store.CheckpointRow;
import com.xc.luckysheet.server.store.OperationRow;
import com.xc.luckysheet.server.store.OutboxRow;
import com.xc.luckysheet.server.store.WorkbookRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;

@Service
public class WorkbookOperationService {
    private static final long CHECKPOINT_OPERATION_LIMIT = 50;
    private static final long CHECKPOINT_BYTES_LIMIT = 512_000;
    public static final String SYSTEM_RESTORE_ACTOR = "system:workbook-restore";

    private final WorkbookStore store;
    private final AccessControlService access;
    private final MutationDescriptorRegistry registry;
    private final ObjectMapper mapper;
    private final AuditRecorder auditRecorder;
    private final CoordinationProperties coordination;
    private final WorkbookDataBlockPublicationGuard dataBlockPublication;
    /**
     * H2 runs the browser service as one JVM.  Keep commit, checkpoint and
     * restore writes for one workbook in one local critical section so a
     * checkpoint cannot race the operation replay entity update.  The
     * database lock remains authoritative for future instances.
     */
    private final ConcurrentHashMap<String, Object> workbookLocks = new ConcurrentHashMap<>();

    public WorkbookOperationService(
            WorkbookStore store,
            AccessControlService access,
            MutationDescriptorRegistry registry,
            ObjectMapper mapper,
            AuditRecorder auditRecorder,
            CoordinationProperties coordination,
            WorkbookDataBlockPublicationGuard dataBlockPublication
    ) {
        this.store = store;
        this.access = access;
        this.registry = registry;
        this.mapper = mapper;
        this.auditRecorder = auditRecorder;
        this.coordination = coordination;
        this.dataBlockPublication = dataBlockPublication;
    }

    public CommitResult operationResult(String unitId, String operationId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        OperationRow row = store.findOperation(operationId).orElseThrow(() -> ServiceException.notFound("Operation not committed"));
        if (!row.unitId().equals(unitId) || !row.actorSubject().equals(actor)) throw ServiceException.forbidden("Operation belongs to another subject");
        return new CommitResult(readCommittedHistoryRow(row), false);
    }

    public WorkbookSnapshotResponse readSnapshot(String unitId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookRow row = requireWorkbook(unitId);
        JsonNode snapshot = currentSnapshot(row);
        if (snapshot.isObject()) ((ObjectNode) snapshot).put("name", row.name());
        String json = writeJson(snapshot);
        return response(unitId, snapshot, row.revision(), checksum(json));
    }

    @Transactional
    public CommitResult commit(String routeUnitId, OperationEnvelope operation, String actor) {
        synchronized (workbookLock(routeUnitId)) {
            try {
                if (operation == null) throw ServiceException.validation("Operation is required");
                return commitInternal(routeUnitId, operation, actor);
            } catch (ServiceException error) {
                auditRecorder.rejected(operation == null ? null : operation.operationId(), routeUnitId, actor, "OPERATION_COMMIT", error.getMessage());
                throw error;
            }
        }
    }

    private CommitResult commitInternal(String routeUnitId, OperationEnvelope operation, String actor) {
        if (!routeUnitId.equals(operation.unitId())) throw ServiceException.validation("Operation unitId does not match route");
        WorkbookAclRole actorRole = access.require(routeUnitId, actor, WorkbookAclRole.VIEWER);
        WorkbookRow row = store.findForUpdate(routeUnitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + routeUnitId));
        if (row.lifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash and cannot accept operations");

        OperationRow existing = store.findOperation(operation.operationId()).orElse(null);
        if (existing != null) {
            if (!existing.actorSubject().equals(actor) || !existing.unitId().equals(routeUnitId)) {
                throw ServiceException.forbidden("Operation belongs to another subject");
            }
            CommittedOperationEnvelope committed = readCommittedHistoryRow(existing);
            List<OperationMutation> originalMutations = committed.mutations().stream()
                    .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList();
            // Server-owned timestamps/ranges are deliberately excluded from request identity.
            if (!committed.clientSessionId().equals(operation.clientSessionId())
                    || committed.clientSequence() != operation.clientSequence()
                    || committed.baseRevision() != operation.baseRevision()
                    || !originalMutations.equals(operation.mutations())
                    || !java.util.Objects.equals(committed.intent(), operation.intent())) {
                throw new ServiceException("OPERATION_ID_REUSED", 409,
                        "Operation " + operation.operationId() + " was already committed with different content; retain the draft and create a new operation");
            }
            return new CommitResult(committed, false);
        }
        OperationRow sequenceExisting = store.findOperationBySequence(routeUnitId, actor, operation.clientSessionId(), operation.clientSequence()).orElse(null);
        if (sequenceExisting != null && !sequenceExisting.operationId().equals(operation.operationId())) {
            throw ServiceException.conflict("clientSequence was already committed");
        }
        CommittedOperationEnvelope undoTarget = validateIntent(routeUnitId, operation, actor, row);
        if (operation.baseRevision() > row.revision()) {
            throw ServiceException.conflict("Base revision is ahead of the server; reload before submitting");
        }
        if (registry.requiresExactBase(operation.mutations()) && row.revision() != operation.baseRevision()) {
            throw ServiceException.conflict("Revision conflict; rebase against revision " + row.revision());
        }

        JsonNode next = currentSnapshot(row);
        List<CommittedOperationMutation> committedMutations = new ArrayList<>();
        for (OperationMutation mutation : operation.mutations()) {
            MutationPreparation prepared = registry.prepare(next, mutation, actorRole);
            boolean ownedSnapshotCommit = registry.usesOwnedSnapshotCommit(prepared, actorRole);
            JsonNode protectionPreimage = ownedSnapshotCommit
                    ? registry.captureProtectionPreimageForOwnedCommit(next, prepared, actorRole)
                    : next;
            WorkbookDataBlockPublicationGuard.PreviousBlockReferences previousBlockReferences = ownedSnapshotCommit
                    ? dataBlockPublication.capturePreviousBlockReferences(next)
                    : null;
            var application = registry.applyPreparedCommit(next, mutation, prepared, actorRole);
            JsonNode candidate = application.snapshot();
            StructuralPatch inversePatch = inverseStructuralPatch(mutation, undoTarget);
            if (inversePatch != null) candidate = registry.applyStructuralPatchOnOwnedSnapshot(candidate, inversePatch);
            StructuralPatch committedPatch = MutationDescriptorRegistry.mergeStructuralPatches(mutation.id(), application.structuralPatch(), inversePatch);
            if (isStructuralPatchMutation(mutation.id()) && committedPatch == null) {
                throw ServiceException.unavailable("STRUCTURAL_PATCH_UNAVAILABLE: structural mutation did not produce server-owned reference facts");
            }
            List<RangeRef> committedRanges = registry.committedRanges(protectionPreimage, prepared, actorRole, committedPatch);
            List<RangeRef> structuralImpactRanges = registry.structuralImpactRanges(committedPatch);
            if (ownedSnapshotCommit) {
                dataBlockPublication.requireNewReferences(routeUnitId, candidate, previousBlockReferences);
            } else {
                dataBlockPublication.requireNewReferences(routeUnitId, next, candidate);
            }
            next = candidate;
            committedMutations.add(CommittedOperationMutation.from(mutation, committedRanges, structuralImpactRanges, committedPatch));
        }

        if (operation.baseRevision() < row.revision()) {
            for (OperationRow intervening : contiguousOperationRowsBetween(routeUnitId, operation.baseRevision(), row.revision())) {
                var interveningMutations = readCommittedHistoryRow(intervening).mutations();
                if (registry.requiresExactBase(interveningMutations.stream()
                        .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList())) {
                    throw ServiceException.conflict("Workbook structure changed after base revision; reload and review the draft");
                }
                for (var previous : interveningMutations) {
                    for (var proposed : committedMutations) {
                        for (var left : allAffectedRanges(previous)) for (var right : allAffectedRanges(proposed)) {
                            if (left.sheetId().equals(right.sheetId()) && left.startRow() <= right.endRow() && left.endRow() >= right.startRow()
                                    && left.startColumn() <= right.endColumn() && left.endColumn() >= right.startColumn()) {
                                throw ServiceException.conflict("Affected cells changed after base revision; reload and review the draft");
                            }
                        }
                    }
                }
            }
        }

        long nextRevision = row.revision() + 1;
        Instant committedAt = Instant.now();
        CommittedOperationEnvelope committed = CommittedOperationEnvelope.from(operation, actor, nextRevision, committedAt, committedMutations);
        String envelopeJson = writeJson(committed);
        store.insertOperation(new OperationRow(operation.operationId(), routeUnitId, nextRevision, actor, operation.clientSessionId(), operation.clientSequence(), operation.baseRevision(), envelopeJson, committedAt));
        enqueueRevisionEvent(routeUnitId, operation.operationId(), nextRevision, envelopeJson, committedAt);
        String canonicalName = next.path("name").asText(row.name()).trim();
        store.updateWorkbookRevisionAndName(routeUnitId, nextRevision, canonicalName, committedAt);
        if (shouldCheckpoint(row, operation, envelopeJson)) {
            String nextJson = writeJson(next);
            store.updateWorkbook(routeUnitId, nextRevision, nextJson, nextRevision, committedAt);
            store.insertCheckpoint(routeUnitId, nextRevision, nextJson, checksum(nextJson), committedAt);
        }
        audit(operation.operationId(), routeUnitId, actor, "OPERATION_COMMIT", "ACCEPTED", null, mapper.createObjectNode().put("revision", nextRevision));
        return new CommitResult(committed, true);
    }

    private CommittedOperationEnvelope validateIntent(String unitId, OperationEnvelope operation, String actor, WorkbookRow row) {
        OperationIntent intent = operation.intent();
        if (intent == null) return null;
        if (!OperationIntent.UNDO.equals(intent.type())) {
            throw ServiceException.validation("Unsupported operation intent");
        }
        if (intent.targetOperationId().equals(operation.operationId())) {
            throw ServiceException.conflict("Undo operation cannot target itself");
        }
        OperationRow targetRow = store.findOperation(intent.targetOperationId()).orElseThrow(
                () -> ServiceException.conflict("Undo target operation is not committed"));
        if (!unitId.equals(targetRow.unitId())) {
            throw ServiceException.forbidden("Undo target belongs to another workbook");
        }
        if (!actor.equals(targetRow.actorSubject())) {
            throw ServiceException.forbidden("Undo target belongs to another subject");
        }
        if (intent.targetBaseRevision() != targetRow.baseRevision()) {
            throw ServiceException.conflict("Undo target base revision does not match the committed operation");
        }
        if (operation.baseRevision() != row.revision()) {
            throw ServiceException.conflict("Undo requires the current workbook revision " + row.revision());
        }
        CommittedOperationEnvelope target = readCommittedHistoryRow(targetRow);
        if (target.mutations().stream().anyMatch(mutation -> mutation.structuralPatch() != null)
                && targetRow.revision() != row.revision()) {
            throw ServiceException.conflict("Structural undo requires the target operation to be the current workbook revision");
        }
        JsonNode structuralUndoPreimage = target.mutations().stream()
                .anyMatch(mutation -> "sheetTable.update".equals(mutation.id()))
                ? snapshotAtRevision(row, targetRow.revision() - 1)
                : null;
        validateStructuralUndoMutations(operation, target, structuralUndoPreimage);
        return target;
    }

    static void validateStructuralUndoMutations(
            OperationEnvelope operation,
            CommittedOperationEnvelope target,
            JsonNode structuralUndoPreimage
    ) {
        boolean[] matchedInverseMutations = new boolean[operation.mutations().size()];
        for (CommittedOperationMutation original : target.mutations()) {
            if (!isStructuralPatchMutation(original.id())) continue;
            if (original.structuralPatch() == null) {
                throw ServiceException.conflict("STRUCTURAL_PATCH_UNAVAILABLE: structural undo target has no server-derived patch");
            }
            List<Integer> matches = new ArrayList<>();
            for (int index = 0; index < operation.mutations().size(); index++) {
                OperationMutation inverse = operation.mutations().get(index);
                if (matchedInverseMutations[index] || !original.sheetId().equals(inverse.sheetId())
                        || !isInverseStructuralMutation(original, inverse)) continue;
                if ("sheetTable.update".equals(original.id())) {
                    String tableId = original.params().path("id").asText("");
                    JsonNode previousTable = findSheetTable(structuralUndoPreimage, original.sheetId(), tableId);
                    if (tableId.isBlank() || previousTable == null || !previousTable.equals(inverse.params())) continue;
                }
                matches.add(index);
            }
            if (matches.size() != 1) {
                throw ServiceException.conflict("Structural undo must contain exactly one matching inverse mutation");
            }
            matchedInverseMutations[matches.get(0)] = true;
        }
        for (int index = 0; index < operation.mutations().size(); index++) {
            if (isStructuralPatchMutation(operation.mutations().get(index).id()) && !matchedInverseMutations[index]) {
                throw ServiceException.conflict("Structural undo contains an unmatched structural mutation");
            }
        }
        boolean targetContainsNonStructuralMutations = target.mutations().stream()
                .anyMatch(mutation -> !isStructuralPatchMutation(mutation.id()));
        if (!targetContainsNonStructuralMutations) {
            for (int index = 0; index < operation.mutations().size(); index++) {
                if (!matchedInverseMutations[index]) {
                    throw ServiceException.conflict("Structural undo contains a mutation outside its target operation");
                }
            }
        }
    }

    private static boolean isStructuralPatchMutation(String mutationId) {
        return GeneratedWorkbookContract.STRUCTURAL_PATCH_MUTATIONS.contains(mutationId);
    }

    private static JsonNode findSheetTable(JsonNode snapshot, String sheetId, String tableId) {
        if (snapshot == null || tableId.isBlank()) return null;
        JsonNode match = null;
        int matches = 0;
        JsonNode sheets = snapshot.path("sheets");
        if (!sheets.isArray()) return null;
        for (JsonNode sheet : sheets) {
            if (!sheetId.equals(sheet.path("id").asText())) continue;
            JsonNode tables = sheet.path("sheetTables");
            if (!tables.isArray()) continue;
            for (JsonNode table : tables) {
                if (!tableId.equals(table.path("id").asText())) continue;
                match = table;
                matches++;
            }
        }
        return matches == 1 ? match : null;
    }

    public static StructuralPatch inverseStructuralPatch(OperationMutation inverse, CommittedOperationEnvelope target) {
        if (target == null) return null;
        List<CommittedOperationMutation> matches = target.mutations().stream()
                .filter(original -> isInverseStructuralMutation(original, inverse)
                        && original.sheetId().equals(inverse.sheetId()))
                .toList();
        if (matches.isEmpty()) return null;
        if (matches.size() != 1) {
            throw ServiceException.conflict("Structural undo target has ambiguous inverse mutations");
        }
        CommittedOperationMutation original = matches.get(0);
        if (original.structuralPatch() == null) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_UNAVAILABLE: structural undo target has no server-derived patch");
        }
        if (!original.id().equals(original.structuralPatch().mutationId())) {
            throw ServiceException.conflict("STRUCTURAL_PATCH_MISMATCH: undo target mutation id does not match its patch");
        }
        return original.structuralPatch().inverse(inverse.id());
    }

    private static boolean isInverseStructuralMutation(CommittedOperationMutation original, OperationMutation inverse) {
        String originalId = original.id();
        String inverseId = inverse.id();
        if (isInverseAxisMutation(originalId, inverseId)) {
            return sameAxisRange(original.params(), inverse.params());
        }
        if ("range.move".equals(originalId) && "range.move".equals(inverseId)) {
            return sameMoveRanges(original.params(), inverse.params());
        }
        if ("rows.permuted".equals(originalId) && "rows.permuted".equals(inverseId)) {
            return sameRowPermutationInverse(original.params(), inverse.params());
        }
        if ("sheetTable.update".equals(originalId) && "sheetTable.update".equals(inverseId)) {
            String tableId = original.params().path("id").asText("");
            return !tableId.isBlank() && tableId.equals(inverse.params().path("id").asText());
        }
        boolean originalCellShiftRestore = originalId.endsWith(".restore");
        boolean inverseCellShiftRestore = inverseId.endsWith(".restore");
        if (originalCellShiftRestore == inverseCellShiftRestore) return false;
        JsonNode originalSpec = originalCellShiftRestore ? original.params().path("spec") : original.params();
        JsonNode inverseSpec = inverseCellShiftRestore ? inverse.params().path("spec") : inverse.params();
        String originalOperation = cellShiftOperation(originalId, originalSpec);
        String inverseOperation = cellShiftOperation(inverseId, inverseSpec);
        return originalOperation != null && originalOperation.equals(inverseOperation)
                && originalSpec.path("operation").asText().equals(inverseSpec.path("operation").asText())
                && originalSpec.path("axis").asText().equals(inverseSpec.path("axis").asText())
                && originalSpec.path("range").equals(inverseSpec.path("range"));
    }

    private static String cellShiftOperation(String mutationId, JsonNode spec) {
        String expected = mutationId.startsWith("cells.inserted") ? "insert"
                : mutationId.startsWith("cells.deleted") ? "delete" : null;
        String actual = spec.path("operation").asText("");
        return expected != null && expected.equals(actual) ? actual : null;
    }

    static boolean sameRowPermutationInverse(JsonNode original, JsonNode inverse) {
        JsonNode range = original.path("range");
        JsonNode inverseRange = inverse.path("range");
        JsonNode sourceRows = original.path("sourceRows");
        JsonNode inverseRows = inverse.path("sourceRows");
        JsonNode affectedColumnEnd = original.path("affectedColumnEnd");
        JsonNode inverseAffectedColumnEnd = inverse.path("affectedColumnEnd");
        if (!range.isObject() || !range.equals(inverseRange)
                || !affectedColumnEnd.isIntegralNumber() || !affectedColumnEnd.equals(inverseAffectedColumnEnd)
                || !range.path("startRow").isIntegralNumber() || !range.path("startRow").canConvertToInt()
                || !sourceRows.isArray() || !inverseRows.isArray()
                || sourceRows.size() == 0 || sourceRows.size() != inverseRows.size()) return false;
        int startRow = range.path("startRow").intValue();
        int[] expectedInverse = new int[sourceRows.size()];
        boolean[] seen = new boolean[sourceRows.size()];
        for (int targetOffset = 0; targetOffset < sourceRows.size(); targetOffset++) {
            JsonNode source = sourceRows.get(targetOffset);
            if (!source.isIntegralNumber() || !source.canConvertToInt()) return false;
            int sourceOffset = source.intValue() - startRow;
            if (sourceOffset < 0 || sourceOffset >= expectedInverse.length || seen[sourceOffset]) return false;
            seen[sourceOffset] = true;
            expectedInverse[sourceOffset] = startRow + targetOffset;
        }
        for (int index = 0; index < expectedInverse.length; index++) {
            JsonNode actual = inverseRows.get(index);
            if (!actual.isIntegralNumber() || !actual.canConvertToInt() || actual.intValue() != expectedInverse[index]) return false;
        }
        return true;
    }

    private static boolean sameMoveRanges(JsonNode original, JsonNode inverse) {
        JsonNode source = original.path("sourceRange");
        JsonNode targetOrigin = original.path("targetOrigin");
        JsonNode inverseSource = inverse.path("sourceRange");
        JsonNode inverseOrigin = inverse.path("targetOrigin");
        for (String coordinate : List.of("startRow", "endRow", "startColumn", "endColumn")) {
            if (!source.path(coordinate).isIntegralNumber() || !inverseSource.path(coordinate).isIntegralNumber()) return false;
        }
        if (!source.path("sheetId").isTextual()
                || !source.path("sheetId").equals(inverseSource.path("sheetId"))
                || !targetOrigin.path("row").isIntegralNumber()
                || !targetOrigin.path("column").isIntegralNumber()
                || !inverseOrigin.path("row").isIntegralNumber()
                || !inverseOrigin.path("column").isIntegralNumber()) return false;
        long startRow = source.path("startRow").asLong();
        long endRow = source.path("endRow").asLong();
        long startColumn = source.path("startColumn").asLong();
        long endColumn = source.path("endColumn").asLong();
        long targetRow = targetOrigin.path("row").asLong();
        long targetColumn = targetOrigin.path("column").asLong();
        return inverseOrigin.path("row").asLong() == startRow
                && inverseOrigin.path("column").asLong() == startColumn
                && inverseSource.path("startRow").asLong() == targetRow
                && inverseSource.path("endRow").asLong() == targetRow + endRow - startRow
                && inverseSource.path("startColumn").asLong() == targetColumn
                && inverseSource.path("endColumn").asLong() == targetColumn + endColumn - startColumn;
    }

    private static boolean isInverseAxisMutation(String originalId, String inverseId) {
        return switch (originalId) {
            case "rows.inserted" -> "rows.deleted".equals(inverseId);
            case "rows.deleted" -> "rows.inserted".equals(inverseId);
            case "columns.inserted" -> "columns.deleted".equals(inverseId);
            case "columns.deleted" -> "columns.inserted".equals(inverseId);
            default -> false;
        };
    }

    private static boolean sameAxisRange(JsonNode original, JsonNode inverse) {
        return original.path("at").isIntegralNumber() && inverse.path("at").isIntegralNumber()
                && original.path("count").isIntegralNumber() && inverse.path("count").isIntegralNumber()
                && original.path("at").asInt() == inverse.path("at").asInt()
                && original.path("count").asInt() == inverse.path("count").asInt();
    }

    private static List<RangeRef> allAffectedRanges(CommittedOperationMutation mutation) {
        List<RangeRef> ranges = new ArrayList<>(mutation.affectedRanges());
        ranges.addAll(mutation.structuralImpactRanges());
        return ranges;
    }

    public CursorPage<RevisionRecord> revisions(String unitId, String actor, long beforeRevision, int limit, String nextCursor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        List<RevisionRecord> items = store.listOperationsBefore(unitId, beforeRevision, limit).stream().map(this::revisionRecord).toList();
        String next = items.size() == limit ? Long.toString(items.get(items.size() - 1).revision()) : null;
        return new CursorPage<>(items, next);
    }

    public WorkbookSnapshotResponse readRevision(String unitId, long revision, String actor) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        WorkbookRow current = requireWorkbook(unitId);
        if (revision < 0 || revision > current.revision()) throw ServiceException.notFound("Revision not found: " + revision);
        JsonNode snapshot = snapshotAtRevision(current, revision);
        String json = writeJson(snapshot);
        return response(unitId, snapshot, revision, checksum(json));
    }

    @Transactional
    public CheckpointResponse checkpoint(String unitId, String actor) {
        synchronized (workbookLock(unitId)) {
            access.require(unitId, actor, WorkbookAclRole.EDITOR);
            WorkbookRow row = store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
            if (row.lifecycle() != WorkbookLifecycle.ACTIVE) throw ServiceException.trashed("Workbook is in trash and cannot be checkpointed");
            if (row.snapshotRevision() == row.revision()) {
                JsonNode snapshot = currentSnapshot(row);
                String canonicalJson = writeJson(snapshot);
                return new CheckpointResponse(unitId, row.revision(), checksum(canonicalJson), false);
            }
            JsonNode snapshot = currentSnapshot(row);
            String json = writeJson(snapshot);
            Instant now = Instant.now();
            store.updateWorkbook(unitId, row.revision(), json, row.revision(), now);
            store.insertCheckpoint(unitId, row.revision(), json, checksum(json), now);
            return new CheckpointResponse(unitId, row.revision(), checksum(json), true);
        }
    }

    @Transactional
    public RestoreResult restore(String unitId, RestoreRequest request, String actor) {
        synchronized (workbookLock(unitId)) {
            access.require(unitId, actor, WorkbookAclRole.OWNER);
            WorkbookRow row = store.findForUpdate(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
            if (request.targetRevision() > row.revision()) throw ServiceException.notFound("Revision not found: " + request.targetRevision());
            JsonNode target = snapshotAtRevision(row, request.targetRevision());
            WorkbookSnapshotValidator.requireCanonical(target, unitId);
            dataBlockPublication.requireSnapshot(unitId, target);
            registry.require("workbook.restore", true);
            long revision = row.revision() + 1;
            Instant now = Instant.now();
            String operationId = UUID.randomUUID().toString();
            ObjectNode params = mapper.createObjectNode()
                    .put("serverGenerated", true)
                    .put("targetRevision", request.targetRevision())
                    .put("reason", request.reason());
            params.set("snapshot", target.deepCopy());
            List<CommittedOperationMutation> mutations = fullWorkbookRestoreMutation(target, params);
            List<OperationMutation> sourceMutations = mutations.stream()
                    .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params())).toList();
            OperationEnvelope source = new OperationEnvelope("system", OperationEnvelope.SCHEMA, operationId, unitId, revision, row.revision(),
                    sourceMutations, now);
            CommittedOperationEnvelope committed = CommittedOperationEnvelope.system(source, actor, revision, now, mutations);
            String json = writeJson(target);
            String envelopeJson = writeJson(committed);
            // System restores must not share an authenticated browser's client-sequence namespace.
            store.insertOperation(new OperationRow(operationId, unitId, revision, SYSTEM_RESTORE_ACTOR, "system", revision, row.revision(), envelopeJson, now));
            enqueueRevisionEvent(unitId, operationId, revision, envelopeJson, now);
            store.updateWorkbook(unitId, revision, json, revision, now);
            store.insertCheckpoint(unitId, revision, json, checksum(json), now);
            audit(operationId, unitId, actor, "SNAPSHOT_RESTORE", "ACCEPTED", request.reason(), mapper.createObjectNode().put("targetRevision", request.targetRevision()));
            return new RestoreResult(committed, response(unitId, target, revision, checksum(json)));
        }
    }

    private Object workbookLock(String unitId) {
        if (unitId == null || unitId.isBlank()) return this;
        return workbookLocks.computeIfAbsent(unitId, ignored -> new Object());
    }

    public List<AclEntry> acl(String unitId, String actor) {
        return access.list(unitId, actor);
    }

    /** Read-only projection; it never accepts a browser-declared actor or role. */
    public WorkbookAccessProjection accessProjection(String unitId, String actor) {
        return new WorkbookAccessProjection(unitId, access.currentRole(unitId, actor));
    }

    public AclEntry grantAcl(String unitId, String actor, String target, WorkbookAclRole role) {
        return access.grant(unitId, actor, target, role);
    }

    public void revokeAcl(String unitId, String actor, String target) {
        access.revoke(unitId, actor, target);
    }

    public List<AuditRecord> audit(String unitId, String actor, int limit) {
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        requireWorkbook(unitId);
        return store.listAudit(unitId, limit);
    }

    private List<CommittedOperationMutation> fullWorkbookRestoreMutation(JsonNode target, ObjectNode params) {
        List<CommittedOperationMutation> mutations = new ArrayList<>();
        JsonNode sheets = target.path("sheets");
        if (sheets.isArray()) {
            for (JsonNode sheet : sheets) {
                String sheetId = sheet.path("id").asText("unknown");
                mutations.add(new CommittedOperationMutation("workbook.restore", sheetId, params, List.of(new RangeRef(sheetId, 0, 1_048_575, 0, 16_383))));
            }
        }
        if (mutations.isEmpty()) mutations.add(new CommittedOperationMutation("workbook.restore", "workbook", params, List.of()));
        return mutations;
    }

    private RevisionRecord revisionRecord(OperationRow row) {
        CommittedOperationEnvelope operation = readCommittedHistoryRow(row);
        return new RevisionRecord(row.operationId(), row.revision(), row.committedAt(), operation);
    }

    private CommittedOperationEnvelope readCommitted(OperationRow row) {
        try {
            return mapper.readValue(row.envelopeJson(), CommittedOperationEnvelope.class);
        } catch (Exception error) {
            throw new ServiceException("STORAGE_CORRUPT", 409,
                    "Stored operation envelope is invalid at revision " + row.revision() + "; restore a verified backup");
        }
    }

    private JsonNode currentSnapshot(WorkbookRow row) {
        CheckpointRow checkpoint = store.findCheckpoint(row.unitId(), row.snapshotRevision())
                .orElseThrow(() -> new ServiceException("STORAGE_CORRUPT", 409, "Checkpoint missing for " + row.unitId() + "; restore a verified backup"));
        verifyCheckpoint(checkpoint);
        if (!checksum(row.snapshotJson()).equals(checkpoint.checksum())) {
            throw new ServiceException("STORAGE_CORRUPT", 409, "Snapshot/checkpoint mismatch for " + row.unitId() + "; restore a verified backup");
        }
        JsonNode snapshot = WorkbookSnapshotValidator.requireCanonical(readJson(row.snapshotJson()), row.unitId());
        if (row.snapshotRevision() == row.revision()) return snapshot;
        for (OperationRow operation : contiguousOperationRowsBetween(row.unitId(), row.snapshotRevision(), row.revision())) {
            snapshot = applyCommittedEnvelope(snapshot, readCommittedHistoryRow(operation));
        }
        return snapshot;
    }

    private JsonNode snapshotAtRevision(WorkbookRow current, long targetRevision) {
        if (targetRevision == current.revision()) return currentSnapshot(current);
        CheckpointRow checkpoint = store.findCheckpoint(current.unitId(), targetRevision)
                .orElseGet(() -> store.findLatestCheckpointAtOrBefore(current.unitId(), targetRevision).orElseThrow(() -> ServiceException.notFound("Snapshot checkpoint not found")));
        verifyCheckpoint(checkpoint);
        JsonNode snapshot = WorkbookSnapshotValidator.requireCanonical(readJson(checkpoint.snapshotJson()), current.unitId());
        if (checkpoint.revision() == targetRevision) return snapshot;
        for (OperationRow operation : contiguousOperationRowsBetween(current.unitId(), checkpoint.revision(), targetRevision)) {
            CommittedOperationEnvelope committed = readCommittedHistoryRow(operation);
            if (committed.mutations().stream().anyMatch(mutation -> "workbook.restore".equals(mutation.id()))) {
                throw ServiceException.conflict("Restore checkpoint is missing for revision " + committed.revision());
            }
            snapshot = applyCommittedEnvelope(snapshot, committed);
        }
        return snapshot;
    }

    private List<OperationRow> contiguousOperationRowsBetween(String unitId, long afterRevision, long throughRevision) {
        if (afterRevision < 0 || throughRevision < afterRevision) {
            throw new ServiceException("STORAGE_CORRUPT", 409,
                    "Invalid operation history range for " + unitId + "; restore a verified backup");
        }
        List<OperationRow> rows = store.listOperationsBetween(unitId, afterRevision, throughRevision);
        long previousRevision = afterRevision;
        for (OperationRow row : rows) {
            long expectedRevision = previousRevision + 1;
            if (!unitId.equals(row.unitId()) || row.revision() != expectedRevision) {
                throw new ServiceException("STORAGE_CORRUPT", 409,
                        "Operation history gap or duplicate for " + unitId + ": expected revision "
                                + expectedRevision + " but found " + row.revision() + "; restore a verified backup");
            }
            previousRevision = row.revision();
        }
        if (previousRevision != throughRevision) {
            throw new ServiceException("STORAGE_CORRUPT", 409,
                    "Operation history for " + unitId + " ends at revision " + previousRevision
                            + " before required revision " + throughRevision + "; restore a verified backup");
        }
        return rows;
    }

    private CommittedOperationEnvelope readCommittedHistoryRow(OperationRow row) {
        CommittedOperationEnvelope committed = readCommitted(row);
        boolean actorMatches = switch (committed.origin()) {
            case CLIENT -> row.actorSubject().equals(committed.actorId());
            case SYSTEM -> SYSTEM_RESTORE_ACTOR.equals(row.actorSubject());
        };
        if (!row.operationId().equals(committed.operationId())
                || !row.unitId().equals(committed.unitId())
                || row.revision() != committed.revision()
                || !actorMatches
                || !row.clientSessionId().equals(committed.clientSessionId())
                || row.clientSequence() != committed.clientSequence()
                || row.baseRevision() != committed.baseRevision()) {
            throw new ServiceException("STORAGE_CORRUPT", 409,
                    "Operation row and envelope identity disagree at revision " + row.revision()
                            + "; restore a verified backup");
        }
        return committed;
    }

    private JsonNode applyCommittedEnvelope(JsonNode snapshot, CommittedOperationEnvelope committed) {
        CommittedOperationEnvelope undoTarget = null;
        if (committed.intent() != null) {
            OperationRow targetRow = store.findOperation(committed.intent().targetOperationId()).orElseThrow(
                    () -> new ServiceException("STORAGE_CORRUPT", 409, "Committed undo target is missing from the operation log"));
            if (!targetRow.unitId().equals(committed.unitId())
                    || !targetRow.actorSubject().equals(committed.actorId())
                    || targetRow.revision() >= committed.revision()) {
                throw new ServiceException("STORAGE_CORRUPT", 409, "Committed undo target identity or revision is invalid");
            }
            undoTarget = readCommittedHistoryRow(targetRow);
            if (!undoTarget.operationId().equals(committed.intent().targetOperationId())) {
                throw new ServiceException("STORAGE_CORRUPT", 409, "Committed undo target operation identity is inconsistent");
            }
        }
        List<StructuralPatch> inversePatches = new ArrayList<>(committed.mutations().size());
        for (CommittedOperationMutation mutation : committed.mutations()) {
            inversePatches.add(inverseStructuralPatch(
                    new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params()),
                    undoTarget));
        }
        return registry.applyCommittedMutations(snapshot, committed.mutations(), inversePatches);
    }

    private WorkbookRow requireWorkbook(String unitId) {
        return store.find(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found: " + unitId));
    }

    private void verifyCheckpoint(CheckpointRow checkpoint) {
        if (!checksum(checkpoint.snapshotJson()).equals(checkpoint.checksum())) {
            throw new ServiceException("STORAGE_CORRUPT", 409, "Checksum mismatch for " + checkpoint.unitId() + " revision " + checkpoint.revision() + "; restore a verified backup");
        }
    }

    private boolean shouldCheckpoint(WorkbookRow row, OperationEnvelope operation, String envelopeJson) {
        if (row.snapshotRevision() == row.revision()) return false;
        List<OperationRow> sinceCheckpoint = store.listOperationsBetween(row.unitId(), row.snapshotRevision(), row.revision());
        long operationCount = sinceCheckpoint.size();
        long bytes = sinceCheckpoint.stream().mapToLong(entry -> entry.envelopeJson().getBytes(StandardCharsets.UTF_8).length).sum();
        return operationCount + 1 >= CHECKPOINT_OPERATION_LIMIT || bytes + envelopeJson.getBytes(StandardCharsets.UTF_8).length >= CHECKPOINT_BYTES_LIMIT;
    }

    private void audit(String operationId, String unitId, String actor, String eventType, String outcome, String reason, JsonNode details) {
        store.insertAudit(new AuditRecord(UUID.randomUUID(), operationId, unitId, actor, eventType, outcome, reason, details, Instant.now()));
    }

    private void enqueueRevisionEvent(String unitId, String operationId, long revision, String envelopeJson, Instant createdAt) {
        if (!coordination.redisEnabled()) return;
        store.insertOutbox(new OutboxRow(UUID.randomUUID(), unitId, operationId, revision, envelopeJson, createdAt, 0));
    }

    private WorkbookSnapshotResponse response(String unitId, JsonNode snapshot, long revision, String checksum) {
        return new WorkbookSnapshotResponse(unitId, snapshot.deepCopy(), revision, checksum);
    }

    private JsonNode readJson(String json) {
        try {
            return mapper.readTree(json);
        } catch (Exception error) {
            throw new IllegalStateException("Stored snapshot JSON is invalid", error);
        }
    }

    private String writeJson(Object json) {
        try {
            return mapper.writeValueAsString(json);
        } catch (Exception error) {
            throw new IllegalStateException("Unable to serialize workbook JSON", error);
        }
    }

    private String checksum(String json) {
        try {
            return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(json.getBytes(StandardCharsets.UTF_8)));
        } catch (Exception error) {
            throw new IllegalStateException("Unable to checksum workbook snapshot", error);
        }
    }

    public record CommitResult(CommittedOperationEnvelope operation, boolean committed) {
    }

    public record RestoreResult(CommittedOperationEnvelope operation, WorkbookSnapshotResponse snapshot) {
    }
}
