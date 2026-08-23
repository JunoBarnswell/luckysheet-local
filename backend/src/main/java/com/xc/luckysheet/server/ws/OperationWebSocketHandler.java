package com.xc.luckysheet.server.ws;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.OperationEnvelope;
import com.xc.luckysheet.server.contract.WorkbookSnapshotResponse;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.coordination.EphemeralCoordinationService;
import com.xc.luckysheet.server.coordination.EphemeralEvent;
import com.xc.luckysheet.server.coordination.WebSocketSessionRegistry;
import com.xc.luckysheet.server.service.AccessControlService;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.ServiceException;
import com.xc.luckysheet.server.service.WorkbookOperationService;
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
    private final WorkbookOperationService operations;
    private final AccessControlService access;
    private final WebSocketSessionRegistry sessions;
    private final EphemeralCoordinationService ephemeral;

    public OperationWebSocketHandler(
            ObjectMapper mapper,
            WorkbookOperationService operations,
            AccessControlService access,
            WebSocketSessionRegistry sessions,
            EphemeralCoordinationService ephemeral
    ) {
        this.mapper = mapper;
        this.operations = operations;
        this.access = access;
        this.sessions = sessions;
        this.ephemeral = ephemeral;
    }

    @Override
    public void handleTextMessage(WebSocketSession session, TextMessage message) {
        String operationId = "unknown";
        try {
            Principal principal = session.getPrincipal();
            String actor = ActorIdentity.subject(principal);
            JsonNode root = mapper.readTree(message.getPayload());
            if (!root.isObject() || root.path("type").asText("").isBlank()) throw ServiceException.validation("Message type is required");
            String type = root.path("type").asText();
            switch (type) {
                case "snapshot.request" -> handleSnapshot(session, actor, root);
                case "changeset.submit" -> {
                    JsonNode payload = root.get("payload");
                    if (payload == null) throw ServiceException.validation("changeset.submit payload is required");
                    OperationEnvelope operation = mapper.treeToValue(payload, OperationEnvelope.class);
                    operationId = operation.operationId().toString();
                    WorkbookOperationService.CommitResult result = operations.commit(operation.unitId(), operation, actor);
                    sessions.join(operation.unitId(), session);
                    send(session, ack(result.operation()));
                    sessions.broadcastRevision(result.operation(), session);
                }
                case "presence.updated", "cursor.updated" -> handleTransient(session, actor, root, type);
                default -> throw ServiceException.validation("Unsupported client collaboration message: " + type);
            }
        } catch (Exception error) {
            sendQuietly(session, reject(operationId, error));
        }
    }

    @Override
    public void afterConnectionClosed(WebSocketSession session, CloseStatus status) {
        String unitId = unitId(session);
        if (unitId != null) {
            sessions.leave(unitId, session);
            EphemeralEvent event = ephemeral.offline(unitId, ActorIdentity.subject(session.getPrincipal()), session.getId());
            sessions.broadcastEphemeral(event, session);
        }
    }

    private void handleSnapshot(WebSocketSession session, String actor, JsonNode root) {
        String unitId = parseUnitId(root.path("unitId").asText(""));
        WorkbookSnapshotResponse snapshot = operations.readSnapshot(unitId, actor);
        sessions.join(unitId, session);
        ObjectNode response = mapper.createObjectNode().put("type", "snapshot.response");
        response.set("payload", mapper.valueToTree(snapshot));
        send(session, response);
    }

    private void handleTransient(WebSocketSession session, String actor, JsonNode root, String type) {
        if (root.has("actorId")) throw ServiceException.validation("actorId is server-owned");
        String unitId = parseUnitId(root.path("unitId").asText(""));
        access.require(unitId, actor, WorkbookAclRole.VIEWER);
        sessions.join(unitId, session);
        EphemeralEvent event = ephemeral.updated(type, unitId, actor, session.getId(), root.get("state"));
        sessions.broadcastEphemeral(event, session);
    }

    private ObjectNode ack(CommittedOperationEnvelope operation) {
        return mapper.createObjectNode().put("type", "changeset.ack")
                .put("operationId", operation.operationId().toString()).put("revision", operation.revision());
    }

    private ObjectNode reject(String operationId, Exception error) {
        String code = error instanceof ServiceException serviceError ? serviceError.code() : "VALIDATION_ERROR";
        String message = error.getMessage() == null ? "Message rejected" : error.getMessage();
        ObjectNode result = mapper.createObjectNode().put("type", "changeset.reject").put("operationId", operationId);
        result.set("error", mapper.createObjectNode().put("code", code).put("message", message));
        return result;
    }

    private void send(WebSocketSession session, ObjectNode message) {
        try {
            session.sendMessage(new TextMessage(mapper.writeValueAsString(message)));
        } catch (IOException error) {
            throw new IllegalStateException("Unable to send collaboration message", error);
        }
    }

    private void sendQuietly(WebSocketSession session, TextMessage message) {
        try {
            if (session.isOpen()) session.sendMessage(message);
        } catch (IOException ignored) {
            // The socket is already closing; no second error can be delivered.
        }
    }

    private void sendQuietly(WebSocketSession session, ObjectNode message) {
        try {
            if (session.isOpen()) session.sendMessage(new TextMessage(mapper.writeValueAsString(message)));
        } catch (IOException ignored) {
            // The socket is already closing; no second error can be delivered.
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
