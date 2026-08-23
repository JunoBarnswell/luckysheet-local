package com.xc.luckysheet.server.ws;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.coordination.EphemeralCoordinationService;
import com.xc.luckysheet.server.coordination.EphemeralEvent;
import com.xc.luckysheet.server.coordination.WebSocketSessionRegistry;
import com.xc.luckysheet.server.service.AccessControlService;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.ServiceException;
import org.springframework.stereotype.Component;
import org.springframework.web.socket.CloseStatus;
import org.springframework.web.socket.TextMessage;
import org.springframework.web.socket.WebSocketSession;
import org.springframework.web.socket.handler.TextWebSocketHandler;

import java.io.IOException;
import java.security.Principal;

@Component
public class OperationWebSocketHandler extends TextWebSocketHandler {
    private final ObjectMapper mapper;
    private final AccessControlService access;
    private final WebSocketSessionRegistry sessions;
    private final EphemeralCoordinationService ephemeral;

    public OperationWebSocketHandler(
            ObjectMapper mapper,
            AccessControlService access,
            WebSocketSessionRegistry sessions,
            EphemeralCoordinationService ephemeral
    ) {
        this.mapper = mapper;
        this.access = access;
        this.sessions = sessions;
        this.ephemeral = ephemeral;
    }

    @Override
    public void handleTextMessage(WebSocketSession session, TextMessage message) {
        try {
            Principal principal = session.getPrincipal();
            String actor = ActorIdentity.subject(principal);
            JsonNode root = mapper.readTree(message.getPayload());
            if (!root.isObject() || root.path("type").asText("").isBlank()) throw ServiceException.validation("Message type is required");
            String type = root.path("type").asText();
            switch (type) {
                case "presence.updated", "cursor.updated" -> handleTransient(session, actor, root, type);
                default -> throw ServiceException.validation("WebSocket accepts presence/cursor updates only; use REST for workbook operations");
            }
        } catch (Exception error) {
            closeForProtocolViolation(session);
        }
    }

    @Override
    public void afterConnectionClosed(WebSocketSession session, CloseStatus status) {
        String unitId = unitId(session);
        if (unitId != null) {
            sessions.leave(unitId, session);
            try {
                EphemeralEvent event = ephemeral.offline(unitId, ActorIdentity.subject(session.getPrincipal()), session.getId());
                sessions.broadcastEphemeral(event, session);
            } catch (ServiceException ignored) {
                // A rejected handshake has no actor to announce as offline.
            }
        }
    }

    private void handleTransient(WebSocketSession session, String actor, JsonNode root, String type) {
        if (root.has("actorId")) throw ServiceException.validation("actorId is server-owned");
        String unitId = parseUnitId(root.path("unitId").asText(""));
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        sessions.join(unitId, session);
        EphemeralEvent event = ephemeral.updated(type, unitId, actor, session.getId(), root.get("state"));
        sessions.broadcastEphemeral(event, session);
    }

    private void closeForProtocolViolation(WebSocketSession session) {
        try {
            if (session.isOpen()) session.close(CloseStatus.POLICY_VIOLATION);
        } catch (IOException ignored) {
            // The peer is already closing; there is no error response channel.
        }
    }

    private String unitId(WebSocketSession session) {
        return sessions.unitId(session);
    }

    private String parseUnitId(String value) {
        if (value == null || value.isBlank() || value.length() > 200) throw ServiceException.validation("unitId is required");
        return value;
    }
}
