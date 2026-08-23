package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.CommittedOperationMutation;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.RangeRef;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.AccessControlService;
import com.xc.luckysheet.server.service.ServiceException;
import org.junit.jupiter.api.Test;
import org.springframework.web.socket.CloseStatus;
import org.springframework.web.socket.WebSocketSession;

import java.security.Principal;
import java.time.Instant;
import java.util.List;
import java.util.Map;

import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.doThrow;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class WebSocketSessionRegistryTest {
    @Test
    void revokedOrUnauthorizedSessionIsClosedBeforeReceivingRemoteRevision() throws Exception {
        AccessControlService access = mock(AccessControlService.class);
        WebSocketSessionRegistry registry = new WebSocketSessionRegistry(new ObjectMapper().findAndRegisterModules(), access);
        WebSocketSession session = mock(WebSocketSession.class);
        Principal principal = () -> "editor-1";
        when(session.getAttributes()).thenReturn(new java.util.concurrent.ConcurrentHashMap<>(Map.of()));
        when(session.getPrincipal()).thenReturn(principal);
        when(session.isOpen()).thenReturn(true);
        doThrow(ServiceException.forbidden("Workbook access denied"))
                .when(access).require("book-1", "editor-1", WorkbookAclRole.VIEWER);
        registry.join("book-1", session);
        CommittedOperationEnvelope operation = new CommittedOperationEnvelope(
                OperationEnvelope.SCHEMA,
                "operation-1",
                "book-1",
                "owner-1",
                1,
                0,
                1,
                List.of(new CommittedOperationMutation("cell.set", "sheet-1", new ObjectMapper().readTree("{\"row\":0,\"column\":0,\"value\":{\"value\":1}}"), List.of(new RangeRef("sheet-1", 0, 0, 0, 0)))),
                Instant.parse("2026-08-23T00:00:00Z"),
                Instant.parse("2026-08-23T00:00:00Z")
        );

        registry.broadcastRevision(operation);

        verify(session).close(eq(CloseStatus.POLICY_VIOLATION));
    }
}
