package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.service.AccessControlService;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.ServiceException;
import com.xc.luckysheet.server.service.RangeAccessResolver;
import com.xc.luckysheet.server.service.RangeAccessService;
import com.xc.luckysheet.server.service.AccessProjectionService;
import org.springframework.beans.factory.annotation.Autowired;
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
    private final RangeAccessService rangeAccess;
    private final AccessProjectionService accessProjection;
    private final Map<String, Set<WebSocketSession>> sessionsByUnit = new ConcurrentHashMap<>();
    private final Map<String, Map<WebSocketSession, String>> calculationBySource = new ConcurrentHashMap<>();
    private final Map<String, Instant> seenRevisionOperations = new ConcurrentHashMap<>();
    private final Map<String, Instant> seenEphemeralEvents = new ConcurrentHashMap<>();

    @Autowired
    public WebSocketSessionRegistry(ObjectMapper mapper, AccessControlService access,
                                    RangeAccessService rangeAccess, AccessProjectionService accessProjection) {
        this.mapper = mapper;
        this.access = access;
        this.rangeAccess = rangeAccess;
        this.accessProjection = accessProjection;
    }

    public void subscribeCalculation(String rootId, java.util.Set<String> sources, WebSocketSession session) {
        unsubscribeCalculation(session);
        for (String source : sources) calculationBySource.computeIfAbsent(source, ignored -> new ConcurrentHashMap<>()).put(session, rootId);
    }

    public void unsubscribeCalculation(WebSocketSession session) {
        calculationBySource.forEach((source, peers) -> {
            peers.remove(session);
            if (peers.isEmpty()) calculationBySource.remove(source, peers);
        });
    }

    private void broadcastCalculationChanged(String source) {
        calculationBySource.getOrDefault(source, Map.of()).forEach((peer, root) -> {
            if (!peer.isOpen()) { unsubscribeCalculation(peer); return; }
            if (!sessionCanRead(root, peer)) { closeRevokedSession(peer); unsubscribeCalculation(peer); return; }
            sendJson(peer, mapper.createObjectNode().put("type", "calculation.changed").put("unitId", root).put("sourceUnitId", source));
        });
    }

    public void broadcastLifecycleChanged(String unitId, String lifecycle) {
        if (!Set.of("active", "trashed", "purged").contains(lifecycle)) throw new IllegalArgumentException("Unknown workbook lifecycle");
        broadcastCalculationChanged(unitId);
        for (WebSocketSession peer : sessionsByUnit.getOrDefault(unitId, Set.of())) {
            if (!peer.isOpen()) continue;
            sendJson(peer, mapper.createObjectNode().put("type", "workbook.lifecycle.changed").put("unitId", unitId).put("lifecycle", lifecycle));
            if (lifecycle.equals("active")) continue;
            unsubscribeCalculation(peer);
            closeRevokedSession(peer);
        }
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
        broadcastCalculationChanged(operation.unitId());
        Set<WebSocketSession> sessions = sessionsByUnit.getOrDefault(operation.unitId(), Set.of());
        for (WebSocketSession peer : sessions) {
            if (peer == origin || !peer.isOpen()) continue;
            if (!sessionCanRead(operation.unitId(), peer)) {
                closeRevokedSession(peer);
                continue;
            }
            ObjectNode message = mapper.createObjectNode().put("type", "revision.created")
                    .put("unitId", operation.unitId()).put("revision", operation.revision()).put("operationId", operation.operationId());
            try {
                RangeAccessResolver resolver = resolverFor(operation.unitId(), peer);
                message.put("accessRevision", resolver.accessRevision());
                if (accessProjection.canDeliver(operation, resolver)) {
                    message.set("payload", mapper.valueToTree(operation));
                } else {
                    message.put("resyncRequired", true);
                }
                sendJson(peer, message);
            } catch (ServiceException error) {
                closeRevokedSession(peer);
            }
        }
    }

    public void broadcastEphemeral(EphemeralEvent event) {
        broadcastEphemeral(event, null);
    }

    public void broadcastEphemeral(EphemeralEvent event, WebSocketSession origin) {
        if (event == null || !markSeen(seenEphemeralEvents, event.eventId())) return;
        Set<WebSocketSession> sessions = sessionsByUnit.getOrDefault(event.unitId(), Set.of());
        for (WebSocketSession peer : sessions) {
            if (peer == origin || !peer.isOpen()) continue;
            if (!sessionCanRead(event.unitId(), peer)) {
                closeRevokedSession(peer);
                continue;
            }
            try {
                RangeAccessResolver resolver = resolverFor(event.unitId(), peer);
                if (!accessProjection.canDeliverEphemeral(event.state(), resolver)) continue;
                ObjectNode message = mapper.createObjectNode()
                        .put("type", event.type().replace(".updated", ".broadcast"))
                        .put("unitId", event.unitId())
                        .put("actorId", event.actorId())
                        .put("accessRevision", resolver.accessRevision());
                message.set("state", event.state().deepCopy());
                sendJson(peer, message);
            } catch (ServiceException error) {
                closeRevokedSession(peer);
            }
        }
    }

    public void broadcastAccessChanged(String unitId, long accessRevision) {
        broadcastCalculationChanged(unitId);
        ObjectNode message = mapper.createObjectNode().put("type", "access.changed")
                .put("unitId", unitId).put("accessRevision", accessRevision);
        broadcast(unitId, null, message);
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
        for (WebSocketSession peer : sessions) {
            if (peer == origin || !peer.isOpen()) continue;
            if (!sessionCanRead(unitId, peer)) {
                closeRevokedSession(peer);
                continue;
            }
            sendJson(peer, message);
        }
    }

    private RangeAccessResolver resolverFor(String unitId, WebSocketSession session) {
        return rangeAccess.resolver(unitId, ActorIdentity.subject(session.getPrincipal()), ActorIdentity.groups(session.getPrincipal()));
    }

    private void sendJson(WebSocketSession session, ObjectNode message) {
        try {
            sendQuietly(session, new TextMessage(mapper.writeValueAsString(message)));
        } catch (Exception error) {
            throw new IllegalStateException("Unable to encode collaboration message", error);
        }
    }

    /** Re-check persistent ACL/share state before every remote delivery. */
    private boolean sessionCanRead(String unitId, WebSocketSession session) {
        try {
            access.require(unitId, ActorIdentity.subject(session.getPrincipal()), WorkbookRole.VIEWER);
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
