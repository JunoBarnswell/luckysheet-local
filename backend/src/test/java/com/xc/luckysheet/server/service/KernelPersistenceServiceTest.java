package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.persistence.*;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.*;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

class KernelPersistenceServiceTest {
    private final ObjectMapper mapper = new ObjectMapper();
    private final WorkbookManifestEntityRepository manifests = mock(WorkbookManifestEntityRepository.class);
    private final WorkbookPageEntityRepository pages = mock(WorkbookPageEntityRepository.class);
    private final WorkbookKernelHistoryEntityRepository histories = mock(WorkbookKernelHistoryEntityRepository.class);
    private final KernelPersistenceService service = new KernelPersistenceService(manifests, pages, histories, mapper);
    private final Map<Long, WorkbookManifestEntity> checkpoints = new HashMap<>();
    private final Map<String, WorkbookPageEntity> contents = new HashMap<>();
    private final Map<String, WorkbookKernelHistoryEntity> records = new HashMap<>();

    KernelPersistenceServiceTest() {
        when(manifests.findByUnitIdAndRevision(eq("unit"), anyLong())).thenAnswer(invocation -> Optional.ofNullable(checkpoints.get(invocation.getArgument(1))));
        when(manifests.save(any())).thenAnswer(invocation -> { WorkbookManifestEntity row = invocation.getArgument(0); checkpoints.put(row.getRevision(), row); return row; });
        when(pages.findByUnitIdAndChecksum(eq("unit"), anyString())).thenAnswer(invocation -> Optional.ofNullable(contents.get(invocation.getArgument(1))));
        when(pages.findContent(eq("unit"), anyString())).thenAnswer(invocation -> Optional.ofNullable(contents.get(invocation.getArgument(1))).map(row -> new WorkbookPageEntityRepository.PageContent() {
            public String getChecksum() { return row.getChecksum(); }
            public long getByteLength() { return row.getByteLength(); }
            public String getPayloadBase64() { return row.getPayloadBase64(); }
        }));
        when(pages.existsByUnitIdAndChecksum(eq("unit"), anyString())).thenAnswer(invocation -> contents.containsKey(invocation.getArgument(1)));
        doAnswer(invocation -> {
            WorkbookPageEntity row = new WorkbookPageEntity(invocation.getArgument(0), invocation.getArgument(1), invocation.getArgument(2),
                    invocation.getArgument(3), invocation.getArgument(4), invocation.getArgument(5));
            contents.put(row.getChecksum(), row); return null;
        }).when(pages).insertContent(anyString(), anyString(), anyString(), anyLong(), anyString(), any());
        when(pages.saveAll(any())).thenAnswer(invocation -> { Iterable<WorkbookPageEntity> rows = invocation.getArgument(0); List<WorkbookPageEntity> saved = new ArrayList<>(); rows.forEach(row -> { contents.put(row.getChecksum(), row); saved.add(row); }); return saved; });
        when(histories.findByUnitIdAndOperationId(eq("unit"), anyString())).thenAnswer(invocation -> Optional.ofNullable(records.get(invocation.getArgument(1))));
        when(histories.save(any())).thenAnswer(invocation -> { WorkbookKernelHistoryEntity row = invocation.getArgument(0); records.put(mapper.readTree(row.getHistoryJson()).path("operationId").asText(), row); return row; });
    }

    @Test
    void preservesUnchangedContentAcrossRevisionAndReopensOnlyTheManifest() throws Exception {
        ObjectNode initial = page(0, 0, 7);
        publishInitial(initial);
        ObjectNode nextManifest = manifest(1, initial);
        ObjectNode change = change("rename", nextManifest);
        nextManifest.put("name", "renamed");
        ObjectNode metadataBefore = change.withObject("history").putObject("metadataBefore");
        metadataBefore.put("name", "book"); metadataBefore.set("sheets", manifest(0, initial).path("sheets")); metadataBefore.putObject("metadata");
        ObjectNode metadataAfter = change.withObject("history").putObject("metadataAfter");
        metadataAfter.put("name", "renamed"); metadataAfter.set("sheets", nextManifest.path("sheets")); metadataAfter.putObject("metadata");
        service.publish(change, "unit", 1);
        assertJsonEquals(initial, service.readPage("unit", 1, "sheet", 0, 0));
        assertJsonEquals(initial, service.readPage("unit", 0, "sheet", 0, 0));
        assertEquals(1, contents.size());
        KernelHostClient kernel = mock(KernelHostClient.class);
        service.reopen("unit", 1, kernel);
        verify(kernel).call(eq("open"), argThat(params -> params.has("manifest") && !params.has("pages")));
        verifyNoMoreInteractions(kernel);
    }

    @Test
    void immutableHistoryReplaysChangedPayloadAndPreservesBeforeBytes() throws Exception {
        ObjectNode initial = page(0, 0, 7); publishInitial(initial);
        ObjectNode updated = page(0, 1, 9);
        ObjectNode change = change("edit", manifest(1, updated));
        change.withArray("pages").add(updated);
        ObjectNode delta = change.withObject("history").withArray("pageDeltas").addObject();
        delta.putObject("key").put("sheetId", "sheet").put("pageRow", 0).put("pageColumn", 0);
        delta.set("before", descriptor(initial)); delta.set("after", descriptor(updated));
        service.publish(change, "unit", 1);
        assertJsonEquals(change, service.readChangeSet("unit", "edit"));
        assertJsonEquals(change.path("history"), service.historyForUndo("unit", "edit"));
        assertJsonEquals(initial, service.readPage("unit", 0, "sheet", 0, 0));
        assertJsonEquals(updated, service.readPage("unit", 1, "sheet", 0, 0));
    }

    @Test
    void rejectsCorruptedBytesBeforeWritingAnyCheckpoint() throws Exception {
        ObjectNode invalid = page(0, 0, 7); invalid.put("payloadBase64", Base64.getEncoder().encodeToString(new byte[invalid.path("byteLength").asInt()]));
        KernelHostException error = assertThrows(KernelHostException.class, () -> publishInitial(invalid));
        assertEquals("PAGE_CHECKSUM_INVALID", error.code());
        verify(manifests, never()).save(any()); verify(pages, never()).saveAll(any()); verify(histories, never()).save(any());
    }

    @Test
    void rejectsHistoryWithForgedBeforeDescriptor() throws Exception {
        ObjectNode initial = page(0, 0, 7); publishInitial(initial);
        ObjectNode updated = page(0, 1, 9);
        ObjectNode change = change("edit", manifest(1, updated)); change.withArray("pages").add(updated);
        ObjectNode delta = change.withObject("history").withArray("pageDeltas").addObject();
        delta.putObject("key").put("sheetId", "sheet").put("pageRow", 0).put("pageColumn", 0);
        delta.putNull("before"); delta.set("after", descriptor(updated));
        KernelHostException error = assertThrows(KernelHostException.class, () -> service.publish(change, "unit", 1));
        assertEquals("HISTORY_INVALID", error.code()); assertFalse(checkpoints.containsKey(1L)); assertEquals(1, contents.size());
    }

    @Test
    void importsAndStagesBoundedRawPageFiles(@TempDir Path root) throws Exception {
        ObjectNode page = page(0, 0, 7);
        Path raw = root.resolve("source.lspg"); Files.write(raw, Base64.getDecoder().decode(page.path("payloadBase64").asText()));
        Path index = root.resolve("pages.json");
        var entries = mapper.createArrayNode(); ObjectNode entry = entries.addObject(); entry.set("descriptor", descriptor(page)); entry.put("fileHandle", raw.toString());
        mapper.writeValue(index.toFile(), entries);
        ObjectNode result = mapper.createObjectNode().put("pagesManifestFile", index.toString()); result.set("manifest", manifest(0, page));
        service.publishImported(result, "unit", 0, root);
        Path staged = service.stagePages("unit", 0, root.resolve("export"));
        JsonNode stagedEntry = mapper.readTree(staged.toFile()).get(0);
        assertJsonEquals(descriptor(page), stagedEntry.path("descriptor"));
        assertArrayEquals(Files.readAllBytes(raw), Files.readAllBytes(Path.of(stagedEntry.path("fileHandle").asText())));
        verify(pages, never()).saveAll(any());
    }

    @Test
    void rejectsImportFilesOutsideTaskRoot(@TempDir Path root) throws Exception {
        Path task = Files.createDirectory(root.resolve("task"));
        ObjectNode page = page(0, 0, 7);
        Path raw = root.resolve("outside.lspg"); Files.write(raw, Base64.getDecoder().decode(page.path("payloadBase64").asText()));
        var entries = mapper.createArrayNode(); ObjectNode entry = entries.addObject(); entry.set("descriptor", descriptor(page)); entry.put("fileHandle", raw.toString());
        Path index = task.resolve("pages.json"); mapper.writeValue(index.toFile(), entries);
        ObjectNode result = mapper.createObjectNode().put("pagesManifestFile", index.toString()); result.set("manifest", manifest(0, page));
        assertEquals("TASK_PATH_INVALID", assertThrows(KernelHostException.class, () -> service.publishImported(result, "unit", 0, task)).code());
        assertTrue(checkpoints.isEmpty()); assertTrue(contents.isEmpty());
    }

    private void assertJsonEquals(JsonNode expected, JsonNode actual) throws Exception {
        // The public contract is JSON, not Jackson's choice of IntNode/LongNode.
        // Serialize both without altering any field or numeric value.
        assertEquals(mapper.readTree(mapper.writeValueAsBytes(expected)), mapper.readTree(mapper.writeValueAsBytes(actual)));
    }

    private void publishInitial(ObjectNode page) {
        ObjectNode initial = mapper.createObjectNode(); initial.set("manifest", manifest(0, page)); initial.putArray("pages").add(page);
        service.publish(initial, "unit", 0);
    }
    private ObjectNode manifest(long revision, ObjectNode page) {
        ObjectNode manifest = mapper.createObjectNode().put("schema", "WorkbookManifest").put("version", 11).put("unitId", "unit").put("name", "book").put("revision", revision);
        manifest.putArray("sheets").addObject().put("sheetId", "sheet").put("name", "Sheet").put("rowCount", 1048576).put("columnCount", 16384).putObject("metadata");
        manifest.putArray("pages").add(descriptor(page)); manifest.putObject("metadata"); return manifest;
    }
    private ObjectNode change(String operationId, ObjectNode manifest) {
        ObjectNode change = mapper.createObjectNode().put("operationId", operationId).put("baseRevision", 0).put("revision", 1);
        change.set("manifest", manifest); change.putArray("pages"); change.putArray("removedPages"); change.putArray("affectedRanges");
        ObjectNode history = change.putObject("history").put("operationId", operationId).put("baseRevision", 0).put("revision", 1);
        history.putArray("pageDeltas"); history.putNull("metadataBefore"); history.putNull("metadataAfter"); return change;
    }
    private ObjectNode descriptor(ObjectNode payload) { ObjectNode result = payload.deepCopy(); result.remove("payloadBase64"); return result; }
    private ObjectNode page(int pageRow, long revision, double value) throws Exception {
        byte[] dictionaries = "{\"text\":[],\"formula\":[],\"metadata\":[],\"errors\":[]}".getBytes(StandardCharsets.UTF_8);
        ByteBuffer bytes = ByteBuffer.allocate(9 + dictionaries.length + 32 * 1024 * 21).order(ByteOrder.LITTLE_ENDIAN);
        bytes.put(new byte[]{'L', 'S', 'P', 'G', 1}).putInt(dictionaries.length).put(dictionaries);
        int offset = bytes.position(); bytes.put(offset, (byte) 4); bytes.putDouble(offset + 1024, value);
        byte[] encoded = bytes.array();
        ObjectNode page = mapper.createObjectNode().put("sheetId", "sheet").put("pageRow", pageRow).put("pageColumn", 0).put("revision", revision)
                .put("checksum", HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(encoded))).put("byteLength", encoded.length).put("cellCount", 1);
        page.putObject("occupiedRange").put("sheetId", "sheet").put("startRow", pageRow * 1024).put("endRow", pageRow * 1024).put("startColumn", 0).put("endColumn", 0);
        page.put("payloadBase64", Base64.getEncoder().encodeToString(encoded)); return page;
    }
}
