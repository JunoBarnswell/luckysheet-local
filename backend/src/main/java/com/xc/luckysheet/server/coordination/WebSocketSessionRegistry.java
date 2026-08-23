package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.AccessControlService;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.ServiceException;
import org.springframework.web.socket.CloseStatus;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;
import org.springframework.web.socket.TextMessage;
import org.springframework.web.socket.WebSocketSession;

import java.io.IOException;
import java.time.Duration;
import java.time.Instant;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

/** Owns local WebSocket membership and receives cross-instance events. */
@Component
public class WebSocketSessionRegistry {
    private static final Logger LOGGER = LoggerFactory.getLogger(WebSocketSessionRegistry.class);
    private static final Duration SEEN_EVENT_RETENTION = Duration.ofMinutes(10);
    private static final int MAX_SEEN_EVENTS = 20_000;

    private final ObjectMapper mapper;
    private final AccessControlService access;
    private final Map<String, Set<WebSocketSession>> sessionsByUnit = new ConcurrentHashMap<>();
    private final Map<String, Instant> seenRevisionOperations = new ConcurrentHashMap<>();
    private final Map<String, Instant> seenEphemeralEvents = new ConcurrentHashMap<>();

    public WebSocketSessionRegistry(ObjectMapper mapper, AccessControlService access) {
        this.mapper = mapper;
        this.access = access;
    }

    public void join(String unitId, WebSocketSession session) {
        String previous = unitId(session);
        if (previous != null && !previous.equals(unitId)) leave(previous, session);
        session.getAttributes().put("unitId", unitId);
        sessionsByUnit.computeIfAbsent(unitId, ignored -> ConcurrentHashMap.newKeySet()).add(session);
    }

    public void leave(String unitId, WebSocketSession session) {
        if (unitId == null) return;
        Set<WebSocketSession> sessions = sessionsByUnit.get(unitId);
        if (sessions == null) return;
        sessions.remove(session);
        if (sessions.isEmpty()) sessionsByUnit.remove(unitId, sessions);
        session.getAttributes().remove("unitId");
    }

    public String unitId(WebSocketSession session) {
        Object value = session.getAttributes().get("unitId");
        return value == null ? null : value.toString();
    }

    public void broadcastRevision(CommittedOperationEnvelope operation) {
        broadcastRevision(operation, null);
    }

    public void broadcastRevision(CommittedOperationEnvelope operation, WebSocketSession origin) {
        if (operation == null || !markSeen(seenRevisionOperations, operation.operationId())) return;
        ObjectNode message = mapper.createObjectNode()
                .put("type", "revision.created")
                .put("revision", operation.revision());
        message.set("payload", mapper.valueToTree(operation));
        broadcast(operation.unitId(), origin, message);
    }

    public void broadcastEphemeral(EphemeralEvent event) {
        broadcastEphemeral(event, null);
    }

    public void broadcastEphemeral(EphemeralEvent event, WebSocketSession origin) {
        if (event == null || !markSeen(seenEphemeralEvents, event.eventId())) return;
        ObjectNode message = mapper.createObjectNode()
                .put("type", event.type().replace(".updated", ".broadcast"))
                .put("unitId", event.unitId())
                .put("actorId", event.actorId());
        message.set("state", event.state().deepCopy());
        broadcast(event.unitId(), origin, message);
    }

    private boolean markSeen(Map<String, Instant> seen, String id) {
        Instant now = Instant.now();
        if (seen.size() > MAX_SEEN_EVENTS) {
            Instant cutoff = now.minus(SEEN_EVENT_RETENTION);
            seen.entrySet().removeIf(entry -> entry.getValue().isBefore(cutoff));
        }
        return seen.putIfAbsent(id, now) == null;
    }

    private void broadcast(String unitId, WebSocketSession origin, ObjectNode message) {
        Set<WebSocketSession> sessions = sessionsByUnit.getOrDefault(unitId, Set.of());
        String json;
        try {
            json = mapper.writeValueAsString(message);
        } catch (Exception error) {
            throw new IllegalStateException("Unable to encode collaboration message", error);
        }
        for (WebSocketSession peer : sessions) {
            if (peer == origin || !peer.isOpen()) continue;
            if (!sessionCanRead(unitId, peer)) {
                closeRevokedSession(peer);
                continue;
            }
            sendQuietly(peer, new TextMessage(json));
        }
    }

    /** Re-check persistent ACL/share state before every remote delivery. */
    private boolean sessionCanRead(String unitId, WebSocketSession session) {
        try {
            access.require(unitId, ActorIdentity.subject(session.getPrincipal()), WorkbookAclRole.VIEWER);
            return true;
        } catch (ServiceException error) {
            return false;
        }
    }

    private void closeRevokedSession(WebSocketSession session) {
        try {
            if (session.isOpen()) session.close(CloseStatus.POLICY_VIOLATION);
        } catch (IOException error) {
            LOGGER.debug("Revoked WebSocket was already closed", error);
        }
    }

    private void sendQuietly(WebSocketSession session, TextMessage message) {
        try {
            if (session.isOpen()) session.sendMessage(message);
        } catch (IOException error) {
            LOGGER.debug("WebSocket peer closed while broadcasting", error);
        }
    }
}
