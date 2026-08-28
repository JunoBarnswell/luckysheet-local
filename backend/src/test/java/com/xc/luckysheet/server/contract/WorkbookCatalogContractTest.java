package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.json.JsonMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class WorkbookCatalogContractTest {
    private final ObjectMapper mapper = JsonMapper.builder().addModule(new JavaTimeModule()).build();

    @Test
    void summaryCarriesActorRoleAndStructuredLocationWithoutLossyPathJoin() throws Exception {
        WorkbookSummary summary = new WorkbookSummary("book-1", "Budget", 3,
                Instant.parse("2026-08-24T00:00:00Z"), WorkbookAclRole.EDITOR, "owner-1", "space-1",
                "folder-1", List.of("团队空间", "财务"), "团队空间", null, WorkbookStorageLocation.REMOTE,
                WorkbookSyncStatus.SYNCED, WorkbookLifecycle.ACTIVE, WorkbookSource.NATIVE, true, null, null);
        String json = mapper.writeValueAsString(summary);
        assertTrue(json.contains("\"role\":\"editor\""));
        assertTrue(json.contains("\"locationPath\":[\"团队空间\",\"财务\"]"));
        assertTrue(json.contains("\"favorite\":true"));
        assertTrue(json.contains("\"storageLocation\":\"remote\""));
        assertTrue(json.contains("\"lifecycle\":\"active\""));
    }

    @Test
    void importedSourceUsesThePublicHyphenatedWireValue() throws Exception {
        assertEquals("\"document-import\"", mapper.writeValueAsString(WorkbookSource.DOCUMENT_IMPORT));
        assertEquals(WorkbookSource.DOCUMENT_IMPORT, mapper.readValue("\"document-import\"", WorkbookSource.class));
        assertEquals("\"personal\"", mapper.writeValueAsString(WorkspaceSpaceType.PERSONAL));
    }
}
