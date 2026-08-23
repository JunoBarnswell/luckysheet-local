package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.contract.AuditRecord;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.time.Instant;
import java.util.UUID;

@Service
public class AuditRecorder {
    private final WorkbookStore store;
    private final ObjectMapper mapper;

    public AuditRecorder(WorkbookStore store, ObjectMapper mapper) {
        this.store = store;
        this.mapper = mapper;
    }

    @Transactional(propagation = Propagation.REQUIRES_NEW)
    public void rejected(String operationId, String unitId, String actor, String eventType, String reason) {
        store.insertAudit(new AuditRecord(
                UUID.randomUUID(), operationId, unitId, actor, eventType, "REJECTED", reason,
                mapper.createObjectNode(), Instant.now()
        ));
    }

    @Transactional(propagation = Propagation.REQUIRES_NEW)
    public void accepted(String operationId, String unitId, String actor, String eventType, String reason, JsonNode details) {
        store.insertAudit(new AuditRecord(
                UUID.randomUUID(), operationId, unitId, actor, eventType, "ACCEPTED", reason,
                details == null ? mapper.createObjectNode() : details.deepCopy(), Instant.now()
        ));
    }
}
