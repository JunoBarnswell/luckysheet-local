package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.fasterxml.jackson.core.JsonToken;
import com.xc.luckysheet.server.persistence.*;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.io.IOException;
import java.time.Instant;
import java.util.*;

/** Atomic v11 checkpoints, content-addressed pages and proven native history. */
@Service
public class KernelPersistenceService {
    private static final List<String> DESCRIPTOR_FIELDS = List.of("sheetId", "pageRow", "pageColumn", "revision", "checksum", "byteLength", "cellCount", "occupiedRange");
    private final WorkbookManifestEntityRepository manifests;
    private final WorkbookPageEntityRepository pages;
    private final WorkbookKernelHistoryEntityRepository histories;
    private final ObjectMapper mapper;

    public KernelPersistenceService(WorkbookManifestEntityRepository manifests, WorkbookPageEntityRepository pages,
                                    WorkbookKernelHistoryEntityRepository histories, ObjectMapper mapper) {
        this.manifests = manifests; this.pages = pages; this.histories = histories; this.mapper = mapper;
    }

    @Transactional
    public void publish(JsonNode changeSet, String unitId, long revision) {
        JsonNode manifest = changeSet.path("manifest");
        Map<PageKey, JsonNode> directory = validateManifest(manifest, unitId, revision);
        if (!changeSet.path("pages").isArray()) throw failure("PAGE_INVALID", "Native result must contain pages", unitId);
        if (manifests.findByUnitIdAndRevision(unitId, revision).isPresent()) throw failure("REVISION_CONFLICT", "Manifest checkpoint already exists", unitId);
        Map<PageKey, JsonNode> before = revision == 0 ? Map.of() : validateManifest(readManifest(unitId, revision - 1), unitId, revision - 1);
        Map<PageKey, JsonNode> changed = new LinkedHashMap<>();
        List<WorkbookPageEntity> pending = new ArrayList<>();
        Set<String> content = new HashSet<>();
        for (JsonNode payload : changeSet.path("pages")) {
            JsonNode descriptor = descriptor(payload);
            PageKey key = key(descriptor);
            if (!sameCanonicalJson(descriptor, directory.get(key)) || changed.putIfAbsent(key, descriptor) != null)
                throw failure("PAGE_DESCRIPTOR_MISMATCH", "Changed page does not match the unique manifest descriptor", key.toString());
            byte[] bytes = validatePayload(payload, descriptor, unitId);
            String checksum = descriptor.path("checksum").asText();
            if (content.add(checksum)) {
                var existing = pages.findByUnitIdAndChecksum(unitId, checksum);
                if (existing.isPresent()) {
                    if (existing.get().getByteLength() != bytes.length || !Arrays.equals(bytes, decodePayload(existing.get().getPayloadBase64(), unitId)))
                        throw failure("PAGE_CHECKSUM_INVALID", "Existing content address contains different bytes", checksum);
                } else pending.add(new WorkbookPageEntity(identity(unitId, checksum), unitId, checksum, bytes.length, payload.path("payloadBase64").asText(), Instant.now()));
            }
        }
        for (var entry : directory.entrySet()) {
            if (!changed.containsKey(entry.getKey()) && !sameCanonicalJson(entry.getValue(), before.get(entry.getKey()))) {
                // Restores reuse content-addressed historical pages. This read
                // verifies durable bytes before publishing the new manifest.
                readDescriptor(unitId, entry.getValue());
                changed.put(entry.getKey(), entry.getValue());
            }
        }
        WorkbookKernelHistoryEntity historyRow = null;
        if (revision > 0) {
            JsonNode history = changeSet.path("history");
            String operationId = requiredText(changeSet, "operationId", unitId);
            if (operationId.length() > 200 || changeSet.path("revision").asLong(-1) != revision || changeSet.path("baseRevision").asLong(-1) != revision - 1)
                throw failure("HISTORY_INVALID", "Changeset revision identity is invalid", operationId);
            validateHistory(history, unitId, operationId, revision - 1, revision, before, directory, manifest);
            if (histories.findByUnitIdAndOperationId(unitId, operationId).isPresent()) throw failure("OPERATION_CONFLICT", "Operation history is immutable", operationId);
            Set<PageKey> expectedChanged = new HashSet<>();
            Set<PageKey> expectedRemoved = new HashSet<>();
            for (JsonNode delta : history.path("pageDeltas")) {
                if (delta.path("after").isNull()) expectedRemoved.add(key(delta.path("key")));
                else expectedChanged.add(key(delta.path("key")));
            }
            if (!expectedChanged.equals(changed.keySet()) || !changeSet.path("removedPages").isArray())
                throw failure("HISTORY_INVALID", "History and changed pages disagree", operationId);
            Set<PageKey> removed = new HashSet<>();
            for (JsonNode item : changeSet.path("removedPages")) if (!removed.add(key(item))) throw failure("HISTORY_INVALID", "Duplicate removed page", operationId);
            if (!removed.equals(expectedRemoved)) throw failure("HISTORY_INVALID", "Removed pages and history disagree", operationId);
            String historyJson = encode(history);
            ObjectNode resultMetadata = ((ObjectNode) changeSet).deepCopy();
            resultMetadata.remove(List.of("manifest", "pages", "history"));
            String resultMetadataJson = encode(resultMetadata);
            historyRow = new WorkbookKernelHistoryEntity(identity(unitId, operationId), unitId, operationId, revision - 1,
                    revision, historyJson, sha256((historyJson + "\n" + resultMetadataJson).getBytes(StandardCharsets.UTF_8)), resultMetadataJson, Instant.now());
        }
        // All contracts are checked before writes. The caller's workbook revision update joins this transaction.
        pages.saveAll(pending);
        String manifestJson = encode(manifest);
        manifests.save(new WorkbookManifestEntity(identity(unitId, Long.toString(revision)), unitId, revision, manifestJson,
                sha256(manifestJson.getBytes(StandardCharsets.UTF_8)), Instant.now()));
        if (historyRow != null) histories.save(historyRow);
    }

    /** Import is the explicit initial checkpoint boundary; raw files never enter an all-pages frame. */
    @Transactional
    public void publishImported(JsonNode result, String unitId, long revision, Path taskRoot) {
        if (revision != 0) throw failure("REVISION_CONFLICT", "Document import requires a new workbook", unitId);
        JsonNode manifest = result.path("manifest");
        Map<PageKey, JsonNode> directory = validateManifest(manifest, unitId, revision);
        requireNewCheckpoint(unitId);
        Set<PageKey> seen = new HashSet<>();
        try {
            Path root = taskRoot.toRealPath();
            Path index = confinedFile(Path.of(requiredText(result, "pagesManifestFile", unitId)), root);
            try (var parser = mapper.getFactory().createParser(index.toFile())) {
                if (parser.nextToken() != JsonToken.START_ARRAY) throw failure("PAGE_INVALID", "Native page index must be an array", index.toString());
                while (parser.nextToken() != JsonToken.END_ARRAY) {
                    if (parser.currentToken() != JsonToken.START_OBJECT) throw failure("PAGE_INVALID", "Native page index entry is invalid", index.toString());
                    JsonNode entry = mapper.readTree(parser);
                    JsonNode descriptor = entry.path("descriptor");
                    PageKey key = key(descriptor);
                    if (!sameCanonicalJson(descriptor, directory.get(key)) || !seen.add(key)) throw failure("PAGE_DESCRIPTOR_MISMATCH", "Import page index differs from manifest", key.toString());
                    Path file = confinedFile(Path.of(requiredText(entry, "fileHandle", unitId)), root);
                    if (Files.size(file) != descriptor.path("byteLength").asLong()) throw failure("PAGE_LENGTH_INVALID", "Import file length differs from descriptor", file.toString());
                    byte[] bytes = Files.readAllBytes(file);
                    ObjectNode payload = ((ObjectNode) descriptor).deepCopy(); payload.put("payloadBase64", Base64.getEncoder().encodeToString(bytes));
                    validatePayload(payload, descriptor, unitId);
                    insertContent(unitId, descriptor, payload.path("payloadBase64").asText());
                }
                if (parser.nextToken() != null) throw failure("PAGE_INVALID", "Native page index has trailing content", index.toString());
            }
            if (!seen.equals(directory.keySet())) throw failure("PAGE_MISSING", "Import page index does not cover every descriptor", unitId);
            saveInitialCheckpoint(unitId, manifest);
        } catch (IOException error) { throw failure("DOCUMENT_IMPORT_FAILED", "Native page files cannot be read: " + error.getMessage(), unitId); }
    }

    /** Exports one immutable page at a time into the native task's file-backed reader directory. */
    @Transactional(readOnly = true)
    public Path stagePages(String unitId, long revision, Path directory) {
        JsonNode manifest = readManifest(unitId, revision);
        try {
            Files.createDirectories(directory);
            Path root = directory.toRealPath();
            Path index = root.resolve("pages.json");
            try (var output = mapper.getFactory().createGenerator(Files.newOutputStream(index, java.nio.file.StandardOpenOption.CREATE_NEW, java.nio.file.StandardOpenOption.WRITE))) {
                output.writeStartArray();
                int pageIndex = 0;
                for (JsonNode descriptor : manifest.path("pages")) {
                    JsonNode payload = readDescriptor(unitId, descriptor);
                    Path file = root.resolve("page-" + pageIndex++ + ".lspg");
                    Files.write(file, decodePayload(payload.path("payloadBase64").asText(), unitId), java.nio.file.StandardOpenOption.CREATE_NEW, java.nio.file.StandardOpenOption.WRITE);
                    ObjectNode entry = mapper.createObjectNode(); entry.set("descriptor", descriptor); entry.put("fileHandle", file.toString());
                    mapper.writeTree(output, entry);
                }
                output.writeEndArray();
            }
            return index;
        } catch (IOException error) { throw failure("DOCUMENT_EXPORT_FAILED", "Native page files cannot be staged: " + error.getMessage(), unitId); }
    }

    /** Copy is an explicit new-workbook boundary; descriptors rebase to revision zero, bytes remain immutable. */
    @Transactional
    public void copyPages(String sourceUnitId, long sourceRevision, String targetUnitId, JsonNode manifest) {
        Map<PageKey, JsonNode> source = validateManifest(readManifest(sourceUnitId, sourceRevision), sourceUnitId, sourceRevision);
        Map<PageKey, JsonNode> target = validateManifest(manifest, targetUnitId, 0);
        requireNewCheckpoint(targetUnitId);
        if (!source.keySet().equals(target.keySet())) throw failure("PAGE_DESCRIPTOR_MISMATCH", "Copy must preserve every source page", targetUnitId);
        for (var entry : source.entrySet()) {
            ObjectNode expected = ((ObjectNode) entry.getValue()).deepCopy(); expected.put("revision", 0);
            if (!sameCanonicalJson(expected, target.get(entry.getKey()))) throw failure("PAGE_DESCRIPTOR_MISMATCH", "Copied descriptor differs from source", entry.getKey().toString());
            JsonNode payload = readDescriptor(sourceUnitId, entry.getValue());
            insertContent(targetUnitId, expected, payload.path("payloadBase64").asText());
        }
        saveInitialCheckpoint(targetUnitId, manifest);
    }

    private void requireNewCheckpoint(String unitId) {
        if (manifests.findByUnitIdAndRevision(unitId, 0).isPresent()) throw failure("REVISION_CONFLICT", "Initial manifest checkpoint already exists", unitId);
    }
    private void saveInitialCheckpoint(String unitId, JsonNode manifest) {
        String json = encode(manifest);
        manifests.save(new WorkbookManifestEntity(identity(unitId, "0"), unitId, 0, json, sha256(json.getBytes(StandardCharsets.UTF_8)), Instant.now()));
    }
    private void insertContent(String unitId, JsonNode descriptor, String payload) {
        String checksum = descriptor.path("checksum").asText();
        // Direct immutable inserts keep the JPA persistence context from retaining all import payloads.
        if (!pages.existsByUnitIdAndChecksum(unitId, checksum)) pages.insertContent(identity(unitId, checksum), unitId, checksum, descriptor.path("byteLength").asLong(), payload, Instant.now());
    }
    private Path confinedFile(Path path, Path root) throws IOException {
        if (!path.isAbsolute()) throw failure("TASK_PATH_INVALID", "Native file handles must be absolute", path.toString());
        Path resolved = path.toRealPath();
        if (!resolved.startsWith(root) || !Files.isRegularFile(resolved)) throw failure("TASK_PATH_INVALID", "Native file handle escapes the task directory", path.toString());
        return resolved;
    }

    @Transactional(readOnly = true)
    public JsonNode readManifest(String unitId, long revision) {
        WorkbookManifestEntity row = manifests.findByUnitIdAndRevision(unitId, revision).orElseThrow(() ->
                failure("KERNEL_STATE_UNAVAILABLE", "Committed v11 manifest revision is missing", unitId + ":" + revision));
        if (row.getManifestVersion() != 11 || !sha256(row.getManifestJson().getBytes(StandardCharsets.UTF_8)).equals(row.getChecksum()))
            throw failure("MANIFEST_INVALID", "Committed manifest checksum or version is invalid", unitId);
        JsonNode manifest = parse(row.getManifestJson(), "MANIFEST_INVALID", unitId);
        validateManifest(manifest, unitId, revision);
        return manifest;
    }

    @Transactional(readOnly = true)
    public JsonNode readPage(String unitId, long revision, String sheetId, int pageRow, int pageColumn) {
        JsonNode descriptor = validateManifest(readManifest(unitId, revision), unitId, revision).get(new PageKey(sheetId, pageRow, pageColumn));
        if (descriptor == null) throw failure("PAGE_NOT_FOUND", "Page is not in the requested manifest", sheetId + ":" + pageRow + ":" + pageColumn);
        return readDescriptor(unitId, descriptor);
    }

    public void reopen(String unitId, long revision, KernelHostClient kernel) {
        kernel.call("open", mapper.createObjectNode().set("manifest", readManifest(unitId, revision)));
    }

    public void loadPage(String unitId, long revision, String sheetId, int pageRow, int pageColumn, KernelHostClient kernel) {
        loadDescriptor(unitId, revision, readPage(unitId, revision, sheetId, pageRow, pageColumn), kernel);
    }

    /** Explicit whole-workbook consumers stream one bounded page frame at a time. */
    public void loadAllPages(String unitId, long revision, KernelHostClient kernel) {
        for (JsonNode descriptor : readManifest(unitId, revision).path("pages")) loadDescriptor(unitId, revision, readDescriptor(unitId, descriptor), kernel);
    }

    @Transactional(readOnly = true)
    public JsonNode historyForUndo(String unitId, String operationId) {
        WorkbookKernelHistoryEntity row = historyRow(unitId, operationId);
        JsonNode history = parse(row.getHistoryJson(), "HISTORY_INVALID", operationId);
        validateHistory(history, unitId, operationId, row.getBaseRevision(), row.getRevision(),
                validateManifest(readManifest(unitId, row.getBaseRevision()), unitId, row.getBaseRevision()),
                validateManifest(readManifest(unitId, row.getRevision()), unitId, row.getRevision()), readManifest(unitId, row.getRevision()));
        return history;
    }

    @Transactional(readOnly = true)
    public JsonNode readChangeSet(String unitId, String operationId) {
        WorkbookKernelHistoryEntity row = historyRow(unitId, operationId);
        ObjectNode result = (ObjectNode) parse(row.getResultMetadataJson(), "HISTORY_INVALID", operationId);
        JsonNode history = historyForUndo(unitId, operationId);
        result.set("history", history);
        result.set("manifest", readManifest(unitId, row.getRevision()));
        ArrayNode resultPages = result.putArray("pages");
        for (JsonNode delta : history.path("pageDeltas")) if (!delta.path("after").isNull()) resultPages.add(readDescriptor(unitId, delta.path("after")));
        return result;
    }

    @Transactional
    public void purge(String unitId) {
        histories.deleteByUnitId(unitId);
        manifests.deleteByUnitId(unitId);
        pages.deleteByUnitId(unitId);
    }

    private WorkbookKernelHistoryEntity historyRow(String unitId, String operationId) {
        WorkbookKernelHistoryEntity row = histories.findByUnitIdAndOperationId(unitId, operationId).orElseThrow(() -> failure("HISTORY_NOT_FOUND", "Committed operation history is missing", operationId));
        if (!sha256((row.getHistoryJson() + "\n" + row.getResultMetadataJson()).getBytes(StandardCharsets.UTF_8)).equals(row.getChecksum())) throw failure("HISTORY_INVALID", "History checksum mismatch", operationId);
        return row;
    }

    private void loadDescriptor(String unitId, long revision, JsonNode payload, KernelHostClient kernel) {
        ObjectNode params = mapper.createObjectNode().put("unitId", unitId).put("revision", revision);
        params.set("page", payload);
        kernel.call("page.load", params);
    }

    private ObjectNode readDescriptor(String unitId, JsonNode descriptor) {
        WorkbookPageEntityRepository.PageContent row = pages.findContent(unitId, descriptor.path("checksum").asText()).orElseThrow(() -> failure("PAGE_MISSING", "Committed page content is missing", key(descriptor).toString()));
        ObjectNode payload = ((ObjectNode) descriptor).deepCopy();
        payload.put("payloadBase64", row.getPayloadBase64());
        if (row.getByteLength() != descriptor.path("byteLength").asLong()) throw failure("PAGE_LENGTH_INVALID", "Stored page length differs from manifest", row.getChecksum());
        validatePayload(payload, descriptor, unitId);
        return payload;
    }

    private Map<PageKey, JsonNode> validateManifest(JsonNode manifest, String unitId, long revision) {
        if (!manifest.isObject() || !"WorkbookManifest".equals(manifest.path("schema").asText()) || manifest.path("version").asInt(-1) != 11
                || !unitId.equals(manifest.path("unitId").asText()) || manifest.path("revision").asLong(-1) != revision
                || !manifest.path("pages").isArray() || !manifest.path("sheets").isArray())
            throw failure("MANIFEST_INVALID", "Canonical v11 manifest identity is invalid", unitId);
        Map<String, JsonNode> sheets = new HashMap<>();
        for (JsonNode sheet : manifest.path("sheets")) {
            String sheetId = requiredText(sheet, "sheetId", unitId);
            if (sheets.putIfAbsent(sheetId, sheet) != null) throw failure("MANIFEST_INVALID", "Duplicate sheet identity", sheetId);
        }
        Map<PageKey, JsonNode> directory = new LinkedHashMap<>();
        for (JsonNode descriptor : manifest.path("pages")) {
            if (!descriptor.equals(descriptor(descriptor))) throw failure("PAGE_DESCRIPTOR_MISMATCH", "Descriptor contains noncanonical fields", unitId);
            PageKey key = key(descriptor);
            JsonNode sheet = sheets.get(key.sheetId());
            long pageRevision = integer(descriptor, "revision", unitId);
            long byteLength = integer(descriptor, "byteLength", unitId);
            long count = integer(descriptor, "cellCount", unitId);
            JsonNode range = descriptor.path("occupiedRange");
            if (sheet == null || pageRevision > revision || byteLength < 1 || byteLength > 1024 * 1024 || count < 1 || count > 32768
                    || !descriptor.path("checksum").asText().matches("[0-9a-f]{64}") || !range.isObject())
                throw failure("PAGE_DESCRIPTOR_MISMATCH", "Descriptor content identity or statistics are invalid", key.toString());
            long firstRow = integer(range, "startRow", unitId), lastRow = integer(range, "endRow", unitId);
            long firstColumn = integer(range, "startColumn", unitId), lastColumn = integer(range, "endColumn", unitId);
            if (!key.sheetId().equals(range.path("sheetId").asText()) || firstRow > lastRow || firstColumn > lastColumn
                    || firstRow / 1024 != key.pageRow() || lastRow / 1024 != key.pageRow() || firstColumn / 32 != key.pageColumn() || lastColumn / 32 != key.pageColumn()
                    || lastRow >= integer(sheet, "rowCount", unitId) || lastColumn >= integer(sheet, "columnCount", unitId)
                    || count > (lastRow - firstRow + 1) * (lastColumn - firstColumn + 1) || directory.putIfAbsent(key, descriptor) != null)
                throw failure("PAGE_STATS_INVALID", "Descriptor occupied range is invalid or duplicated", key.toString());
        }
        return directory;
    }

    private void validateHistory(JsonNode history, String unitId, String operationId, long baseRevision, long revision,
                                 Map<PageKey, JsonNode> before, Map<PageKey, JsonNode> after, JsonNode newManifest) {
        if (!history.isObject() || !operationId.equals(history.path("operationId").asText()) || history.path("baseRevision").asLong(-1) != baseRevision
                || history.path("revision").asLong(-1) != revision || !history.path("pageDeltas").isArray()) throw failure("HISTORY_INVALID", "History revision identity is invalid", operationId);
        Set<PageKey> expected = new HashSet<>(before.keySet()); expected.addAll(after.keySet());
        expected.removeIf(key -> sameCanonicalJson(before.get(key), after.get(key)));
        Set<PageKey> seen = new HashSet<>();
        for (JsonNode delta : history.path("pageDeltas")) {
            PageKey key = key(delta.path("key"));
            if (!seen.add(key) || !sameCanonicalJson(nullable(delta.get("before")), before.get(key)) || !sameCanonicalJson(nullable(delta.get("after")), after.get(key)))
                throw failure("HISTORY_INVALID", "History page references do not match immutable checkpoints", key.toString());
        }
        if (!seen.equals(expected)) throw failure("HISTORY_INVALID", "History must describe every changed page exactly once", operationId);
        JsonNode oldManifest = readManifest(unitId, baseRevision);
        JsonNode metadataBefore = history.path("metadataBefore");
        JsonNode metadataAfter = history.path("metadataAfter");
        ObjectNode actual = mapper.createObjectNode();
        actual.set("name", oldManifest.path("name")); actual.set("sheets", oldManifest.path("sheets")); actual.set("metadata", oldManifest.path("metadata"));
        boolean metadataChanged = !sameCanonicalJson(oldManifest.path("name"), newManifest.path("name"))
                || !sameCanonicalJson(oldManifest.path("sheets"), newManifest.path("sheets")) || !sameCanonicalJson(oldManifest.path("metadata"), newManifest.path("metadata"));
        ObjectNode next = mapper.createObjectNode();
        next.set("name", newManifest.path("name")); next.set("sheets", newManifest.path("sheets")); next.set("metadata", newManifest.path("metadata"));
        if (metadataChanged ? (!sameCanonicalJson(actual, metadataBefore) || !sameCanonicalJson(next, metadataAfter))
                : (!metadataBefore.isNull() || !metadataAfter.isNull()))
            throw failure("HISTORY_INVALID", "History metadata does not match the checkpoint transition", operationId);
    }
    private static JsonNode nullable(JsonNode node) { return node == null || node.isNull() ? null : node; }
    private static boolean sameCanonicalJson(JsonNode left, JsonNode right) {
        if (left == null || right == null) return left == right;
        return left.equals((a, b) -> {
            if (a.isIntegralNumber() && b.isIntegralNumber()) return a.bigIntegerValue().compareTo(b.bigIntegerValue());
            return a.equals(b) ? 0 : 1;
        }, right);
    }
    private ObjectNode descriptor(JsonNode source) {
        ObjectNode descriptor = mapper.createObjectNode();
        for (String field : DESCRIPTOR_FIELDS) {
            if (!source.has(field)) throw failure("PAGE_DESCRIPTOR_MISMATCH", "Missing descriptor field " + field, source.path("sheetId").asText());
            descriptor.set(field, source.get(field));
        }
        return descriptor;
    }
    private byte[] validatePayload(JsonNode payload, JsonNode descriptor, String unitId) {
        byte[] bytes = decodePayload(requiredText(payload, "payloadBase64", unitId), unitId);
        if (bytes.length != descriptor.path("byteLength").asLong(-1)) throw failure("PAGE_LENGTH_INVALID", "Page byte length differs from descriptor", unitId);
        if (!sha256(bytes).equals(descriptor.path("checksum").asText())) throw failure("PAGE_CHECKSUM_INVALID", "Page checksum differs from bytes", unitId);
        return bytes;
    }
    private byte[] decodePayload(String payload, String object) {
        if (payload.length() > 1_398_104) throw failure("PAGE_TOO_LARGE", "Native page exceeds 1 MiB", object);
        try {
            byte[] bytes = Base64.getDecoder().decode(payload);
            if (bytes.length > 1024 * 1024 || !Base64.getEncoder().encodeToString(bytes).equals(payload)) throw failure("PAGE_INVALID", "Noncanonical or oversized page payload", object);
            return bytes;
        } catch (IllegalArgumentException error) { throw failure("PAGE_INVALID", "Page payload is not base64", object); }
    }
    private PageKey key(JsonNode node) {
        long row = integer(node, "pageRow", "page"), column = integer(node, "pageColumn", "page");
        if (row > 1023 || column > 511) throw failure("PAGE_ADDRESS_INVALID", "Page coordinates exceed worksheet limits", "page");
        return new PageKey(requiredText(node, "sheetId", "page"), (int) row, (int) column);
    }
    private static long integer(JsonNode node, String field, String object) {
        JsonNode value = node.path(field);
        if (!value.isIntegralNumber() || !value.canConvertToLong() || value.longValue() < 0) throw failure("MANIFEST_INVALID", field + " must be an unsigned integer", object);
        return value.longValue();
    }
    private static String requiredText(JsonNode node, String field, String object) {
        JsonNode value = node.path(field);
        if (!value.isTextual() || value.textValue().isBlank()) throw failure("MANIFEST_INVALID", field + " must be nonempty text", object);
        return value.textValue();
    }
    private String encode(JsonNode node) {
        try { return mapper.writeValueAsString(node); }
        catch (Exception error) { throw failure("KERNEL_RESPONSE_INVALID", "Native JSON cannot be encoded", "json"); }
    }
    private JsonNode parse(String json, String code, String object) {
        try { return mapper.readTree(json); }
        catch (Exception error) { throw failure(code, "Committed JSON cannot be decoded", object); }
    }
    private static String identity(String unitId, String component) { return sha256((unitId.length() + ":" + unitId + component).getBytes(StandardCharsets.UTF_8)); }
    private static String sha256(byte[] bytes) {
        try { return HexFormat.of().formatHex(java.security.MessageDigest.getInstance("SHA-256").digest(bytes)); }
        catch (java.security.NoSuchAlgorithmException error) { throw new IllegalStateException(error); }
    }
    private static KernelHostException failure(String code, String message, String object) {
        return new KernelHostException(code, message, object, "reject-and-reopen-last-committed-manifest");
    }
    private record PageKey(String sheetId, int pageRow, int pageColumn) { }
}
