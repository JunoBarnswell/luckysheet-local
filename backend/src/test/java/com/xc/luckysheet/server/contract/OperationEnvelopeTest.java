package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.json.JsonMapper;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import org.junit.jupiter.api.Test;

import java.time.Instant;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class OperationEnvelopeTest {
    private final ObjectMapper mapper = JsonMapper.builder().addModule(new JavaTimeModule()).configure(com.fasterxml.jackson.databind.DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, true).build();

    @Test
    void requestContractDoesNotAllowClientActorOrRanges() throws Exception {
        OperationEnvelope operation = new OperationEnvelope(
                OperationEnvelope.SCHEMA,
                "op-1",
                "unit-1",
                1,
                0,
                List.of(new OperationMutation("cell.set", "sheet-1", mapper.readTree("{\"row\":0,\"column\":0,\"value\":{\"value\":\"ok\"}}"))),
                Instant.parse("2026-08-23T00:00:00Z")
        );
        String json = mapper.writeValueAsString(operation);
        assertFalse(json.contains("actorId"));
        assertFalse(json.contains("affectedRanges"));
        assertThrows(Exception.class, () -> mapper.readValue(json.replace("\"clientSequence\":1", "\"actorId\":\"spoof\",\"clientSequence\":1"), OperationEnvelope.class));
    }

    @Test
    void committedContractContainsServerIdentityAndRanges() throws Exception {
        JsonNode params = mapper.readTree("{\"row\":0,\"column\":0,\"value\":{\"value\":\"ok\"}}");
        CommittedOperationEnvelope operation = CommittedOperationEnvelope.from(
                new OperationEnvelope(OperationEnvelope.SCHEMA, "op-1", "unit-1", 1, 0, List.of(new OperationMutation("cell.set", "sheet-1", params)), Instant.now()),
                "subject-1", 1, Instant.now(), List.of(new CommittedOperationMutation("cell.set", "sheet-1", params, List.of(new RangeRef("sheet-1", 0, 0, 0, 0))))
        );
        String json = mapper.writeValueAsString(operation);
        assertTrue(json.contains("subject-1"));
        assertTrue(json.contains("affectedRanges"));
    }
}
