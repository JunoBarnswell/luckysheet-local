package com.xc.luckysheet.server.migration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.OperationMutation;
import com.xc.luckysheet.server.contract.StructuralPatch;
import com.xc.luckysheet.server.contract.WorkbookSnapshotValidator;
import com.xc.luckysheet.server.mutation.MutationDescriptorRegistry;
import com.xc.luckysheet.server.service.WorkbookOperationService;
import org.flywaydb.core.api.migration.BaseJavaMigration;
import org.flywaydb.core.api.migration.Context;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.HexFormat;

/** One-time, fail-closed replay boundary from verified legacy/v1 history to canonical v2 owner deltas. */
public abstract class StructuralPatchV2Migration extends BaseJavaMigration {
    private final ObjectMapper mapper = new ObjectMapper().registerModule(new JavaTimeModule());

    private record WorkbookHead(String unitId, long snapshotRevision, long revision) { }
    private record Checkpoint(String json, String checksum, JsonNode snapshot) { }
    private record StoredOperation(String operationId, String actorSubject, String clientSessionId,
            long clientSequence, long baseRevision, long revision, String envelopeJson) { }

    @Override
    public Integer getChecksum() {
        return 2;
    }

    @Override
    public void migrate(Context context) throws Exception {
        Connection connection = context.getConnection();
        MutationDescriptorRegistry registry = new MutationDescriptorRegistry();
        List<WorkbookHead> workbooks = listWorkbooks(connection);
        validateCurrentSnapshotPairs(connection, workbooks);
        validatePendingOutboxSourcePairs(connection);
        canonicalizeDefinedNameModels(connection);
        for (WorkbookHead workbook : workbooks) {
            migrateWorkbook(connection, registry, workbook);
        }
        migratePendingOutbox(connection);
    }

    private void validateCurrentSnapshotPairs(Connection connection, List<WorkbookHead> workbooks) throws Exception {
        for (WorkbookHead workbook : workbooks) {
            if (workbook.snapshotRevision() < 0 || workbook.revision() < 0
                    || workbook.snapshotRevision() > workbook.revision()) {
                throw failure("WORKBOOK_REVISION_INVALID", workbook.unitId(), "snapshot revision is outside workbook history");
            }
            Checkpoint checkpoint = readCheckpoint(connection, workbook.unitId(), workbook.snapshotRevision());
            String snapshotJson = readWorkbookSnapshot(connection, workbook.unitId());
            if (!checksum(snapshotJson).equals(checkpoint.checksum())
                    || !mapper.readTree(snapshotJson).equals(checkpoint.snapshot())) {
                throw failure("WORKBOOK_CHECKPOINT_MISMATCH", workbook.unitId(), "workbooks.snapshot_json differs from snapshot_revision checkpoint");
            }
        }
    }

    private void validatePendingOutboxSourcePairs(Connection connection) throws Exception {
        try (PreparedStatement query = connection.prepareStatement(
                    "select event_id, unit_id, operation_id, revision, payload_json from coordination_outbox where published_at is null order by unit_id, revision");
             PreparedStatement operationQuery = connection.prepareStatement(
                    "select envelope_json from operation_log where unit_id = ? and operation_id = ? and revision = ?")) {
            try (ResultSet rows = query.executeQuery()) {
                while (rows.next()) {
                    String eventId = rows.getString(1);
                    String unitId = rows.getString(2);
                    String operationId = rows.getString(3);
                    long revision = rows.getLong(4);
                    ObjectNode payload = parseEnvelope(rows.getString(5), unitId, revision);
                    if (!operationId.equals(payload.path("operationId").asText())
                            || payload.path("revision").asLong(-1) != revision) {
                        throw failure("OUTBOX_IDENTITY_MISMATCH", unitId, "pending event identity differs: " + eventId);
                    }
                    operationQuery.setString(1, unitId);
                    operationQuery.setString(2, operationId);
                    operationQuery.setLong(3, revision);
                    try (ResultSet operation = operationQuery.executeQuery()) {
                        if (!operation.next()) {
                            throw failure("OUTBOX_OPERATION_MISSING", unitId, "operation log row is missing at revision " + revision);
                        }
                        JsonNode committed = mapper.readTree(operation.getString(1));
                        if (operation.next()) {
                            throw failure("OUTBOX_OPERATION_AMBIGUOUS", unitId, "operation log row is duplicated at revision " + revision);
                        }
                        if (!payload.equals(committed)) {
                            throw failure("OUTBOX_SOURCE_MISMATCH", unitId,
                                    "pending payload differs from its operation log source at revision " + revision);
                        }
                    }
                }
            }
        }
    }

    private void canonicalizeDefinedNameModels(Connection connection) throws Exception {
        try (PreparedStatement query = connection.prepareStatement(
                    "select unit_id, revision, snapshot_json, checksum from snapshot_checkpoint order by unit_id, revision");
             PreparedStatement update = connection.prepareStatement(
                    "update snapshot_checkpoint set snapshot_json = ?, checksum = ? where unit_id = ? and revision = ?")) {
            try (ResultSet rows = query.executeQuery()) {
                while (rows.next()) {
                    String unitId = rows.getString(1);
                    long revision = rows.getLong(2);
                    String json = rows.getString(3);
                    if (!checksum(json).equals(rows.getString(4))) {
                        throw failure("CHECKPOINT_CHECKSUM", unitId, "checkpoint checksum is invalid at revision " + revision);
                    }
                    ObjectNode snapshot = (ObjectNode) WorkbookSnapshotValidator.requireCanonical(mapper.readTree(json), unitId);
                    if (!ensureDefinedNameModels(snapshot, unitId)) continue;
                    String canonicalJson = mapper.writeValueAsString(snapshot);
                    update.setString(1, canonicalJson);
                    update.setString(2, checksum(canonicalJson));
                    update.setString(3, unitId);
                    update.setLong(4, revision);
                    if (update.executeUpdate() != 1) {
                        throw failure("CHECKPOINT_WRITE_FAILED", unitId, "checkpoint changed at revision " + revision);
                    }
                }
            }
        }
        try (PreparedStatement query = connection.prepareStatement("select unit_id, snapshot_json from workbooks order by unit_id");
             PreparedStatement update = connection.prepareStatement("update workbooks set snapshot_json = ? where unit_id = ?")) {
            try (ResultSet rows = query.executeQuery()) {
                while (rows.next()) {
                    String unitId = rows.getString(1);
                    ObjectNode snapshot = (ObjectNode) WorkbookSnapshotValidator.requireCanonical(
                            mapper.readTree(rows.getString(2)), unitId);
                    if (!ensureDefinedNameModels(snapshot, unitId)) continue;
                    update.setString(1, mapper.writeValueAsString(snapshot));
                    update.setString(2, unitId);
                    if (update.executeUpdate() != 1) {
                        throw failure("WORKBOOK_WRITE_FAILED", unitId, "workbook snapshot changed during migration");
                    }
                }
            }
        }
    }

    private boolean ensureDefinedNameModels(ObjectNode snapshot, String unitId) {
        WorkbookSnapshotValidator.requireCanonical(snapshot, unitId);
        if (snapshot.has("definedNameModels")) return false;
        ArrayNode models = snapshot.putArray("definedNameModels");
        JsonNode projection = snapshot.get("definedNames");
        if (projection != null) {
            projection.fields().forEachRemaining(entry -> models.addObject()
                    .put("name", entry.getKey())
                    .put("formula", entry.getValue().asText())
                    .put("scope", "workbook"));
        }
        WorkbookSnapshotValidator.requireCanonical(snapshot, unitId);
        return true;
    }

    private List<WorkbookHead> listWorkbooks(Connection connection) throws Exception {
        List<WorkbookHead> workbooks = new ArrayList<>();
        try (PreparedStatement query = connection.prepareStatement(
                "select unit_id, snapshot_revision, revision from workbooks order by unit_id");
             ResultSet rows = query.executeQuery()) {
            while (rows.next()) workbooks.add(new WorkbookHead(rows.getString(1), rows.getLong(2), rows.getLong(3)));
        }
        return workbooks;
    }

    private void migrateWorkbook(Connection connection, MutationDescriptorRegistry registry, WorkbookHead workbook) throws Exception {
        if (workbook.snapshotRevision() < 0 || workbook.revision() < 0
                || workbook.snapshotRevision() > workbook.revision()) {
            throw failure("WORKBOOK_REVISION_INVALID", workbook.unitId(), "snapshot revision is outside workbook history");
        }
        Set<Long> checkpointRevisions = checkpointRevisions(connection, workbook);
        if (!checkpointRevisions.contains(0L) || !checkpointRevisions.contains(workbook.snapshotRevision())) {
            throw failure("MISSING_BASELINE", workbook.unitId(), "revision-0 and current snapshot checkpoints are required");
        }

        Checkpoint baseline = readCheckpoint(connection, workbook.unitId(), 0);
        JsonNode current = baseline.snapshot();
        String storedSnapshotJson = readWorkbookSnapshot(connection, workbook.unitId());
        Checkpoint currentCheckpoint = workbook.snapshotRevision() == 0
                ? baseline
                : readCheckpoint(connection, workbook.unitId(), workbook.snapshotRevision());
        if (!checksum(storedSnapshotJson).equals(currentCheckpoint.checksum())
                || !mapper.readTree(storedSnapshotJson).equals(currentCheckpoint.snapshot())) {
            throw failure("WORKBOOK_CHECKPOINT_MISMATCH", workbook.unitId(), "workbooks.snapshot_json differs from snapshot_revision checkpoint");
        }
        if (workbook.snapshotRevision() == 0 && !current.equals(currentCheckpoint.snapshot())) {
            throw failure("BASELINE_MISMATCH", workbook.unitId(), "revision-0 workbook snapshot differs from checkpoint");
        }

        Set<String> undoTargetIds = undoTargetIds(connection, workbook.unitId());
        Set<String> seenOperationIds = new HashSet<>();
        Map<String, CommittedOperationEnvelope> migratedUndoTargets = new java.util.HashMap<>();
        long previousRevision = 0;
        Checkpoint checkpointAtWorkbookSnapshot = workbook.snapshotRevision() == 0 ? currentCheckpoint : null;
        try (PreparedStatement query = connection.prepareStatement(
                    "select operation_id, actor_subject, client_session_id, client_sequence, base_revision, revision, envelope_json "
                            + "from operation_log where unit_id = ? order by revision");
             PreparedStatement update = connection.prepareStatement(
                    "update operation_log set envelope_json = ? where operation_id = ? and unit_id = ? and revision = ?")) {
            query.setString(1, workbook.unitId());
            query.setFetchSize(128);
            try (ResultSet rows = query.executeQuery()) {
                while (rows.next()) {
                    StoredOperation row = new StoredOperation(rows.getString(1), rows.getString(2), rows.getString(3),
                            rows.getLong(4), rows.getLong(5), rows.getLong(6), rows.getString(7));
                    long expectedRevision = previousRevision + 1;
                    if (row.revision() != expectedRevision || row.revision() > workbook.revision()) {
                        throw failure("OPERATION_GAP", workbook.unitId(), "expected revision " + expectedRevision + " but found " + row.revision());
                    }
                    ObjectNode rawEnvelope = parseEnvelope(row.envelopeJson(), workbook.unitId(), row.revision());
                    canonicalizeRestoreSnapshots(rawEnvelope, workbook.unitId(), row.revision());
                    CommittedOperationEnvelope operation = readEnvelopeWithoutPatches(rawEnvelope, workbook.unitId(), row.revision());
                    validateOperationRow(row, operation, workbook.unitId());
                    if (!seenOperationIds.add(operation.operationId())) {
                        throw failure("DUPLICATE_OPERATION", workbook.unitId(), "duplicate operationId " + operation.operationId());
                    }

                    CommittedOperationEnvelope undoTarget = null;
                    if (operation.intent() != null) {
                        undoTarget = migratedUndoTargets.get(operation.intent().targetOperationId());
                        if (undoTarget == null) {
                            throw failure("UNDO_TARGET_MISSING", workbook.unitId(), "target " + operation.intent().targetOperationId()
                                    + " is missing or does not precede revision " + row.revision());
                        }
                        validateUndoTarget(operation, undoTarget, workbook.unitId());
                    }

                    if (isRestore(operation)) {
                        ensureRestoreHasNoStructuralPatch(rawEnvelope, workbook.unitId(), row.revision());
                        Checkpoint restoreCheckpoint = requireCheckpointAt(connection, checkpointRevisions, workbook.unitId(), row.revision());
                        JsonNode embedded = restoreSnapshot(operation, workbook.unitId(), row.revision());
                        if (!embedded.equals(restoreCheckpoint.snapshot())) {
                            throw failure("RESTORE_CHECKPOINT_MISMATCH", workbook.unitId(), "restore envelope differs from checkpoint at revision " + row.revision());
                        }
                        current = restoreCheckpoint.snapshot();
                    } else {
                        List<OperationMutation> mutations = operation.mutations().stream()
                                .map(mutation -> new OperationMutation(mutation.id(), mutation.sheetId(), mutation.params()))
                                .toList();
                        CommittedOperationEnvelope resolvedUndoTarget = undoTarget;
                        List<StructuralPatch> inversePatches = new ArrayList<>(mutations.size());
                        for (OperationMutation mutation : mutations) {
                            inversePatches.add(WorkbookOperationService.inverseStructuralPatch(mutation, resolvedUndoTarget));
                        }
                        MutationDescriptorRegistry.StructuralPatchMigrationReplay replay = registry.replayStructuralPatchesForMigration(
                                current, mutations, inversePatches);
                        rewriteMutationPatches(rawEnvelope, operation, replay.structuralPatches(), registry, workbook.unitId(), row.revision());
                        current = replay.snapshot();
                    }

                    String migratedJson = mapper.writeValueAsString(rawEnvelope);
                    update.setString(1, migratedJson);
                    update.setString(2, row.operationId());
                    update.setString(3, workbook.unitId());
                    update.setLong(4, row.revision());
                    if (update.executeUpdate() != 1) {
                        throw failure("OPERATION_WRITE_FAILED", workbook.unitId(), "operation row changed at revision " + row.revision());
                    }
                    CommittedOperationEnvelope migrated = readEnvelope(rawEnvelope, workbook.unitId(), row.revision());
                    if (undoTargetIds.contains(migrated.operationId())) migratedUndoTargets.put(migrated.operationId(), migrated);

                    if (checkpointRevisions.contains(row.revision())) {
                        Checkpoint checkpoint = readCheckpoint(connection, workbook.unitId(), row.revision());
                        if (!current.equals(checkpoint.snapshot())) {
                            throw failure("HISTORY_CHECKPOINT_MISMATCH", workbook.unitId(), "replay differs from checkpoint at revision " + row.revision());
                        }
                        if (row.revision() == workbook.snapshotRevision()) checkpointAtWorkbookSnapshot = checkpoint;
                    }
                    previousRevision = row.revision();
                }
            }
        }
        if (previousRevision != workbook.revision()) {
            throw failure("OPERATION_HISTORY_TRUNCATED", workbook.unitId(), "history ends at revision " + previousRevision
                    + " before workbook revision " + workbook.revision());
        }
        if (checkpointAtWorkbookSnapshot == null || !currentCheckpoint.checksum().equals(checkpointAtWorkbookSnapshot.checksum())) {
            throw failure("CURRENT_CHECKPOINT_UNVERIFIED", workbook.unitId(), "snapshot_revision checkpoint was not reached by contiguous history");
        }
    }

    private Set<Long> checkpointRevisions(Connection connection, WorkbookHead workbook) throws Exception {
        Set<Long> revisions = new HashSet<>();
        try (PreparedStatement query = connection.prepareStatement(
                "select revision from snapshot_checkpoint where unit_id = ? order by revision")) {
            query.setString(1, workbook.unitId());
            try (ResultSet rows = query.executeQuery()) {
                while (rows.next()) {
                    long revision = rows.getLong(1);
                    if (revision < 0 || revision > workbook.revision() || !revisions.add(revision)) {
                        throw failure("CHECKPOINT_REVISION_INVALID", workbook.unitId(), "invalid checkpoint revision " + revision);
                    }
                }
            }
        }
        return revisions;
    }

    private Checkpoint requireCheckpointAt(Connection connection, Set<Long> revisions, String unitId, long revision) throws Exception {
        if (!revisions.contains(revision)) throw failure("RESTORE_CHECKPOINT_MISSING", unitId, "restore revision " + revision + " has no checkpoint");
        return readCheckpoint(connection, unitId, revision);
    }

    private Checkpoint readCheckpoint(Connection connection, String unitId, long revision) throws Exception {
        try (PreparedStatement query = connection.prepareStatement(
                "select snapshot_json, checksum from snapshot_checkpoint where unit_id = ? and revision = ?")) {
            query.setString(1, unitId);
            query.setLong(2, revision);
            try (ResultSet row = query.executeQuery()) {
                if (!row.next()) throw failure("CHECKPOINT_MISSING", unitId, "checkpoint revision " + revision + " is missing");
                String json = row.getString(1);
                String expectedChecksum = row.getString(2);
                if (row.next()) throw failure("CHECKPOINT_AMBIGUOUS", unitId, "checkpoint revision " + revision + " is duplicated");
                if (!checksum(json).equals(expectedChecksum)) {
                    throw failure("CHECKPOINT_CHECKSUM", unitId, "checkpoint checksum is invalid at revision " + revision);
                }
                JsonNode snapshot = mapper.readTree(json);
                WorkbookSnapshotValidator.requireCanonical(snapshot, unitId);
                return new Checkpoint(json, expectedChecksum, snapshot);
            }
        }
    }

    private String readWorkbookSnapshot(Connection connection, String unitId) throws Exception {
        try (PreparedStatement query = connection.prepareStatement("select snapshot_json from workbooks where unit_id = ?")) {
            query.setString(1, unitId);
            try (ResultSet row = query.executeQuery()) {
                if (!row.next()) throw failure("WORKBOOK_MISSING", unitId, "workbook row disappeared during migration");
                String json = row.getString(1);
                if (row.next()) throw failure("WORKBOOK_AMBIGUOUS", unitId, "workbook identity is duplicated");
                return json;
            }
        }
    }

    private Set<String> undoTargetIds(Connection connection, String unitId) throws Exception {
        Set<String> targets = new HashSet<>();
        try (PreparedStatement query = connection.prepareStatement(
                "select envelope_json from operation_log where unit_id = ? order by revision")) {
            query.setString(1, unitId);
            try (ResultSet rows = query.executeQuery()) {
                while (rows.next()) {
                    ObjectNode envelope = parseEnvelope(rows.getString(1), unitId, -1);
                    JsonNode intent = envelope.get("intent");
                    if (intent == null || intent.isNull()) continue;
                    if (!intent.isObject() || !intent.path("targetOperationId").isTextual()
                            || intent.path("targetOperationId").asText().isBlank()) {
                        throw failure("UNDO_INTENT_INVALID", unitId, "operation intent has no targetOperationId");
                    }
                    targets.add(intent.path("targetOperationId").asText());
                }
            }
        }
        return targets;
    }

    private ObjectNode parseEnvelope(String json, String unitId, long revision) throws Exception {
        JsonNode value = mapper.readTree(json);
        if (!(value instanceof ObjectNode envelope) || !"OperationEnvelope".equals(envelope.path("schema").asText())
                || !envelope.path("unitId").asText().equals(unitId)) {
            throw failure("ENVELOPE_INVALID", unitId, "invalid operation envelope at revision " + revision);
        }
        if (!(envelope.get("mutations") instanceof ArrayNode)) {
            throw failure("ENVELOPE_MUTATIONS_INVALID", unitId, "operation mutations are invalid at revision " + revision);
        }
        return envelope;
    }

    private CommittedOperationEnvelope readEnvelopeWithoutPatches(ObjectNode raw, String unitId, long revision) throws Exception {
        ObjectNode copy = raw.deepCopy();
        ArrayNode mutations = (ArrayNode) copy.get("mutations");
        for (JsonNode entry : mutations) {
            if (!(entry instanceof ObjectNode mutation)) throw failure("MUTATION_INVALID", unitId, "malformed mutation at revision " + revision);
            mutation.remove("structuralPatch");
        }
        try {
            return mapper.treeToValue(copy, CommittedOperationEnvelope.class);
        } catch (Exception error) {
            throw failure("ENVELOPE_CONTRACT_INVALID", unitId, "operation envelope contract is invalid at revision " + revision, error);
        }
    }

    private CommittedOperationEnvelope readEnvelope(ObjectNode raw, String unitId, long revision) throws Exception {
        try {
            return mapper.treeToValue(raw, CommittedOperationEnvelope.class);
        } catch (Exception error) {
            throw failure("V2_ENVELOPE_INVALID", unitId, "rewritten v2 envelope is invalid at revision " + revision, error);
        }
    }

    private void validateOperationRow(StoredOperation row, CommittedOperationEnvelope operation, String unitId) {
        boolean actorMatches = switch (operation.origin()) {
            case CLIENT -> row.actorSubject().equals(operation.actorId());
            case SYSTEM -> WorkbookOperationService.SYSTEM_RESTORE_ACTOR.equals(row.actorSubject());
        };
        if (!row.operationId().equals(operation.operationId()) || !unitId.equals(operation.unitId())
                || row.revision() != operation.revision() || !actorMatches
                || !row.clientSessionId().equals(operation.clientSessionId())
                || row.clientSequence() != operation.clientSequence() || row.baseRevision() != operation.baseRevision()) {
            throw failure("OPERATION_ROW_MISMATCH", unitId, "database identity differs from envelope at revision " + row.revision());
        }
    }

    private void validateUndoTarget(CommittedOperationEnvelope operation, CommittedOperationEnvelope target, String unitId) {
        if (!target.operationId().equals(operation.intent().targetOperationId())
                || !unitId.equals(target.unitId()) || !target.actorId().equals(operation.actorId())
                || target.revision() >= operation.revision()
                || operation.intent().targetBaseRevision() != target.baseRevision()) {
            throw failure("UNDO_TARGET_INVALID", unitId, "undo target identity or revision is invalid at revision " + operation.revision());
        }
    }

    private boolean isRestore(CommittedOperationEnvelope operation) {
        return operation.mutations().stream().anyMatch(mutation -> "workbook.restore".equals(mutation.id()));
    }

    private JsonNode restoreSnapshot(CommittedOperationEnvelope operation, String unitId, long revision) {
        JsonNode expected = null;
        for (CommittedOperationMutation mutation : operation.mutations()) {
            if (!"workbook.restore".equals(mutation.id())) {
                throw failure("RESTORE_MUTATION_MIXED", unitId, "restore operation contains a non-restore mutation at revision " + revision);
            }
            JsonNode snapshot = mutation.params().get("snapshot");
            if (snapshot == null || !snapshot.isObject() || expected != null && !expected.equals(snapshot)) {
                throw failure("RESTORE_SNAPSHOT_INVALID", unitId, "restore mutation snapshots disagree at revision " + revision);
            }
            expected = snapshot;
        }
        if (expected == null) throw failure("RESTORE_SNAPSHOT_MISSING", unitId, "restore snapshot is missing at revision " + revision);
        WorkbookSnapshotValidator.requireCanonical(expected, unitId);
        return expected;
    }

    private void canonicalizeRestoreSnapshots(ObjectNode envelope, String unitId, long revision) {
        ArrayNode mutations = (ArrayNode) envelope.get("mutations");
        for (JsonNode rawMutation : mutations) {
            if (!(rawMutation instanceof ObjectNode mutation) || !"workbook.restore".equals(mutation.path("id").asText())) continue;
            JsonNode rawParams = mutation.get("params");
            if (!(rawParams instanceof ObjectNode params) || !(params.get("snapshot") instanceof ObjectNode snapshot)) {
                throw failure("RESTORE_SNAPSHOT_INVALID", unitId, "restore snapshot is invalid at revision " + revision);
            }
            ensureDefinedNameModels(snapshot, unitId);
        }
    }

    private void ensureRestoreHasNoStructuralPatch(ObjectNode raw, String unitId, long revision) {
        ArrayNode rawMutations = (ArrayNode) raw.get("mutations");
        for (int index = 0; index < rawMutations.size(); index++) {
            ObjectNode mutation = (ObjectNode) rawMutations.get(index);
            JsonNode patch = mutation.get("structuralPatch");
            JsonNode impacts = mutation.get("structuralImpactRanges");
            if (patch != null && !patch.isNull()
                    || impacts != null && !impacts.isNull() && (!(impacts instanceof ArrayNode array) || !array.isEmpty())) {
                throw failure("RESTORE_PATCH_UNEXPECTED", unitId, "restore operation has structural owner state at revision " + revision);
            }
            mutation.remove("structuralPatch");
            mutation.set("structuralImpactRanges", mapper.createArrayNode());
        }
    }

    private boolean containsRestoreMutation(ObjectNode envelope) {
        JsonNode mutations = envelope.get("mutations");
        if (!(mutations instanceof ArrayNode array)) return false;
        for (JsonNode mutation : array) {
            if (mutation.isObject() && "workbook.restore".equals(mutation.path("id").asText())) return true;
        }
        return false;
    }

    private void rewriteMutationPatches(
            ObjectNode rawEnvelope,
            CommittedOperationEnvelope operation,
            List<Optional<StructuralPatch>> patches,
            MutationDescriptorRegistry registry,
            String unitId,
            long revision
    ) {
        ArrayNode rawMutations = (ArrayNode) rawEnvelope.get("mutations");
        if (rawMutations.size() != operation.mutations().size() || patches.size() != operation.mutations().size()) {
            throw failure("MUTATION_COUNT_MISMATCH", unitId, "reducer mutation count differs at revision " + revision);
        }
        for (int index = 0; index < rawMutations.size(); index++) {
            ObjectNode rawMutation = (ObjectNode) rawMutations.get(index);
            Optional<StructuralPatch> derived = patches.get(index);
            JsonNode oldPatch = rawMutation.get("structuralPatch");
            List<com.xc.luckysheet.server.contract.RangeRef> expectedImpact = derived
                    .map(registry::structuralImpactRanges).orElseGet(List::of);
            JsonNode oldImpact = rawMutation.get("structuralImpactRanges");
            JsonNode expectedImpactNode = mapper.valueToTree(expectedImpact);
            if (derived.isEmpty()) {
                if (oldPatch != null && !oldPatch.isNull()) {
                    throw failure("UNEXPECTED_STRUCTURAL_PATCH", unitId, "stored patch has no reducer owner at revision " + revision);
                }
                if (oldImpact != null && !oldImpact.isNull()
                        && (!(oldImpact instanceof ArrayNode ranges) || !ranges.isEmpty())) {
                    throw failure("UNEXPECTED_STRUCTURAL_IMPACT", unitId, "stored impact has no reducer patch at revision " + revision);
                }
                rawMutation.remove("structuralPatch");
                rawMutation.set("structuralImpactRanges", expectedImpactNode);
                continue;
            }
            StructuralPatch patch = derived.orElseThrow();
            if (oldPatch == null || oldPatch.isNull()) {
                // The pre-v1 envelope contract had neither structuralPatch nor structuralImpactRanges.
                // Its verified replay is the migration authority for generating the first canonical patch.
                if (oldImpact != null && !oldImpact.isNull()) {
                    throw failure("STRUCTURAL_PATCH_PARTIAL_LEGACY", unitId,
                            "legacy structural mutation has impact metadata but no v1 patch at revision " + revision);
                }
                rawMutation.set("structuralPatch", mapper.valueToTree(patch));
                rawMutation.set("structuralImpactRanges", expectedImpactNode);
                continue;
            }
            if (oldImpact == null || !oldImpact.equals(expectedImpactNode)) {
                throw failure("STRUCTURAL_IMPACT_MISMATCH", unitId, "stored impact differs from reducer at revision " + revision);
            }
            if (oldPatch == null || !oldPatch.isObject()
                    || !oldPatch.path("version").isIntegralNumber() || oldPatch.path("version").asInt(-1) != 1
                    || !patch.mutationId().equals(oldPatch.path("mutationId").asText())
                    || !oldPatch.path("formulaOwnerDeltas").isArray()
                    || oldPatch.has("definedNameOwnerDeltas")) {
                throw failure("STRUCTURAL_PATCH_V1_INVALID", unitId, "stored v1 patch is invalid at revision " + revision);
            }
            Set<String> fields = new HashSet<>();
            oldPatch.fieldNames().forEachRemaining(fields::add);
            if (!fields.equals(Set.of("version", "mutationId", "formulaOwnerDeltas"))) {
                throw failure("STRUCTURAL_PATCH_V1_FIELDS", unitId, "stored v1 patch has a non-canonical field set at revision " + revision);
            }
            JsonNode expectedFormulaOwners = mapper.valueToTree(patch).get("formulaOwnerDeltas");
            if (!oldPatch.get("formulaOwnerDeltas").equals(expectedFormulaOwners)) {
                throw failure("STRUCTURAL_FORMULA_PATCH_MISMATCH", unitId, "formula owners changed while upgrading revision " + revision);
            }
            rawMutation.set("structuralPatch", mapper.valueToTree(patch));
            rawMutation.set("structuralImpactRanges", expectedImpactNode);
        }
    }

    private void migratePendingOutbox(Connection connection) throws Exception {
        try (PreparedStatement query = connection.prepareStatement(
                    "select event_id, unit_id, operation_id, revision, payload_json from coordination_outbox where published_at is null order by unit_id, revision");
             PreparedStatement operationQuery = connection.prepareStatement(
                    "select envelope_json from operation_log where unit_id = ? and operation_id = ? and revision = ?");
             PreparedStatement update = connection.prepareStatement(
                    "update coordination_outbox set payload_json = ? where event_id = ? and published_at is null")) {
            query.setFetchSize(128);
            try (ResultSet rows = query.executeQuery()) {
                while (rows.next()) {
                    String eventId = rows.getString(1);
                    String unitId = rows.getString(2);
                    String operationId = rows.getString(3);
                    long revision = rows.getLong(4);
                    ObjectNode oldPayload = parseEnvelope(rows.getString(5), unitId, revision);
                    canonicalizeRestoreSnapshots(oldPayload, unitId, revision);
                    if (containsRestoreMutation(oldPayload)) ensureRestoreHasNoStructuralPatch(oldPayload, unitId, revision);
                    if (!operationId.equals(oldPayload.path("operationId").asText())
                            || oldPayload.path("revision").asLong(-1) != revision) {
                        throw failure("OUTBOX_IDENTITY_MISMATCH", unitId, "pending outbox identity differs at revision " + revision);
                    }
                    operationQuery.setString(1, unitId);
                    operationQuery.setString(2, operationId);
                    operationQuery.setLong(3, revision);
                    String committedJson;
                    try (ResultSet operationRow = operationQuery.executeQuery()) {
                        if (!operationRow.next()) throw failure("OUTBOX_OPERATION_MISSING", unitId, "operation log row is missing at revision " + revision);
                        committedJson = operationRow.getString(1);
                        if (operationRow.next()) throw failure("OUTBOX_OPERATION_AMBIGUOUS", unitId, "operation log row is duplicated at revision " + revision);
                    }
                    ObjectNode committed = parseEnvelope(committedJson, unitId, revision);
                    if (!withoutStructuralMetadata(oldPayload).equals(withoutStructuralMetadata(committed))) {
                        throw failure("OUTBOX_ENVELOPE_MISMATCH", unitId, "pending payload differs from operation log at revision " + revision);
                    }
                    CommittedOperationEnvelope canonicalOperation = readEnvelope(committed, unitId, revision);
                    validatePendingOutboxPatches(oldPayload, canonicalOperation, unitId, revision);
                    update.setString(1, committedJson);
                    update.setString(2, eventId);
                    if (update.executeUpdate() != 1) throw failure("OUTBOX_WRITE_FAILED", unitId, "pending event changed: " + eventId);
                }
            }
        }
    }

    private void validatePendingOutboxPatches(
            ObjectNode oldPayload,
            CommittedOperationEnvelope canonicalOperation,
            String unitId,
            long revision
    ) {
        JsonNode rawMutations = oldPayload.get("mutations");
        if (!(rawMutations instanceof ArrayNode oldMutations)
                || oldMutations.size() != canonicalOperation.mutations().size()) {
            throw failure("OUTBOX_MUTATION_COUNT_MISMATCH", unitId, "pending mutation count differs at revision " + revision);
        }
        for (int index = 0; index < oldMutations.size(); index++) {
            JsonNode oldMutation = oldMutations.get(index);
            if (!oldMutation.isObject()) {
                throw failure("OUTBOX_MUTATION_INVALID", unitId, "pending mutation is malformed at revision " + revision);
            }
            JsonNode oldPatch = oldMutation.get("structuralPatch");
            JsonNode oldImpact = oldMutation.get("structuralImpactRanges");
            StructuralPatch canonicalPatch = canonicalOperation.mutations().get(index).structuralPatch();
            if (canonicalPatch == null) {
                if (oldPatch != null && !oldPatch.isNull()) {
                    throw failure("OUTBOX_STRUCTURAL_PATCH_MISMATCH", unitId, "pending payload has an unexpected v1 patch at revision " + revision);
                }
                if (oldImpact != null && !oldImpact.isNull()
                        && (!(oldImpact instanceof ArrayNode ranges) || !ranges.isEmpty())) {
                    throw failure("OUTBOX_STRUCTURAL_IMPACT_MISMATCH", unitId, "pending payload has unexpected structural impact at revision " + revision);
                }
                continue;
            }
            JsonNode expectedImpact = mapper.valueToTree(canonicalOperation.mutations().get(index).structuralImpactRanges());
            if (oldPatch == null || oldPatch.isNull()) {
                // Legacy outbox events did not persist either structural field; the migrated operation log is authoritative.
                if (oldImpact != null && !oldImpact.isNull()) {
                    throw failure("OUTBOX_STRUCTURAL_PATCH_PARTIAL_LEGACY", unitId,
                            "pending legacy mutation has impact metadata but no v1 patch at revision " + revision);
                }
                continue;
            }
            if (oldImpact == null || !oldImpact.equals(expectedImpact)) {
                throw failure("OUTBOX_STRUCTURAL_IMPACT_MISMATCH", unitId, "pending v1 impact differs from operation log at revision " + revision);
            }
            if (oldPatch == null || !oldPatch.isObject()
                    || !oldPatch.path("version").isIntegralNumber() || oldPatch.path("version").asInt(-1) != 1
                    || !canonicalPatch.mutationId().equals(oldPatch.path("mutationId").asText())
                    || !oldPatch.path("formulaOwnerDeltas").isArray()
                    || oldPatch.has("definedNameOwnerDeltas")) {
                throw failure("OUTBOX_STRUCTURAL_PATCH_INVALID", unitId, "pending v1 patch is invalid at revision " + revision);
            }
            Set<String> fields = new HashSet<>();
            oldPatch.fieldNames().forEachRemaining(fields::add);
            if (!fields.equals(Set.of("version", "mutationId", "formulaOwnerDeltas"))
                    || !oldPatch.get("formulaOwnerDeltas").equals(mapper.valueToTree(canonicalPatch.formulaOwnerDeltas()))) {
                throw failure("OUTBOX_STRUCTURAL_PATCH_MISMATCH", unitId, "pending v1 patch differs from operation log at revision " + revision);
            }
        }
    }

    private ObjectNode withoutStructuralMetadata(ObjectNode input) {
        ObjectNode copy = input.deepCopy();
        JsonNode mutations = copy.get("mutations");
        if (mutations instanceof ArrayNode array) {
            for (JsonNode entry : array) {
                if (entry instanceof ObjectNode mutation) {
                    mutation.remove("structuralPatch");
                    mutation.remove("structuralImpactRanges");
                }
            }
        }
        return copy;
    }

    private String checksum(String json) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(json.getBytes(StandardCharsets.UTF_8)));
    }

    private IllegalStateException failure(String code, String unitId, String message) {
        return new IllegalStateException(code + ": " + unitId + ": " + message);
    }

    private IllegalStateException failure(String code, String unitId, String message, Throwable cause) {
        return new IllegalStateException(code + ": " + unitId + ": " + message, cause);
    }
}
