package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.json.JsonMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertEquals;

class WorkbookAclRoleTest {
    private final ObjectMapper mapper = JsonMapper.builder().addModule(new JavaTimeModule()).build();

    @Test
    void roleUsesLowerCaseWireValueAndAcceptsTheBrowserContract() throws Exception {
        assertEquals("\"viewer\"", mapper.writeValueAsString(WorkbookAclRole.VIEWER));
        assertEquals(WorkbookAclRole.EDITOR, mapper.readValue("\"editor\"", WorkbookAclRole.class));
        assertEquals(WorkbookAclRole.COMMENTER, mapper.readValue("\"COMMENTER\"", WorkbookAclRole.class));
    }

    @Test
    void accessProjectionDoesNotExposeJavaEnumCapitalization() throws Exception {
        String json = mapper.writeValueAsString(new WorkbookAccessProjection("book-1", WorkbookAclRole.EDITOR));
        assertEquals("{\"unitId\":\"book-1\",\"role\":\"editor\"}", json);
    }

    @Test
    void guestShareResponseUsesTheSameLowerCaseRoleContract() throws Exception {
        ShareResponse response = new ShareResponse(
                UUID.fromString("00000000-0000-0000-0000-000000000001"),
                "book-1",
                WorkbookAclRole.COMMENTER,
                Instant.parse("2026-08-24T00:00:00Z"),
                null,
                "owner-1",
                Instant.parse("2026-08-23T00:00:00Z"),
                "secret"
        );
        assertEquals(true, mapper.writeValueAsString(response).contains("\"role\":\"commenter\""));
    }
}
