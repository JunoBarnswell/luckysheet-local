package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.json.JsonMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class WorkbookRoleTest {
    private final ObjectMapper mapper = JsonMapper.builder().addModule(new JavaTimeModule()).build();

    @Test
    void roleUsesLowerCaseWireValueAndAcceptsTheBrowserContract() throws Exception {
        assertEquals("\"viewer\"", mapper.writeValueAsString(WorkbookRole.VIEWER));
        assertEquals(WorkbookRole.EDITOR, mapper.readValue("\"editor\"", WorkbookRole.class));
        assertEquals(WorkbookRole.COMMENTER, mapper.readValue("\"commenter\"", WorkbookRole.class));
        assertThrows(Exception.class, () -> mapper.readValue("\"COMMENTER\"", WorkbookRole.class));
        assertThrows(Exception.class, () -> mapper.readValue("\"admin\"", WorkbookRole.class));
    }

    @Test
    void roleHierarchyIsGeneratedAndDistinctFromSystemAdministration() {
        assertEquals(true, WorkbookRole.OWNER.includes(WorkbookRole.EDITOR));
        assertEquals(false, WorkbookRole.VIEWER.includes(WorkbookRole.COMMENTER));
        assertEquals(true, WorkbookRole.COMMENTER.includes(WorkbookRole.VIEWER));
    }

    @Test
    void accessProjectionDoesNotExposeJavaEnumCapitalization() throws Exception {
        String json = mapper.writeValueAsString(new WorkbookAccessProjection("book-1", WorkbookRole.EDITOR, 0, java.util.List.of()));
        assertEquals("{\"unitId\":\"book-1\",\"role\":\"editor\",\"accessRevision\":0,\"regions\":[]}", json);
    }

    @Test
    void guestShareResponseUsesTheSameLowerCaseRoleContract() throws Exception {
        ShareResponse response = new ShareResponse(
                UUID.fromString("00000000-0000-0000-0000-000000000001"),
                "book-1",
                WorkbookRole.COMMENTER,
                Instant.parse("2026-08-24T00:00:00Z"),
                null,
                "owner-1",
                Instant.parse("2026-08-23T00:00:00Z"),
                "secret"
        );
        assertEquals(true, mapper.writeValueAsString(response).contains("\"role\":\"commenter\""));
    }
}
