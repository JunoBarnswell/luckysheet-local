package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.JsonNode;

import java.util.Objects;
import java.util.UUID;

public record EphemeralEvent(
        String eventId,
        String type,
        String unitId,
        String actorId,
        String sessionId,
        JsonNode state
) {
    public EphemeralEvent {
        if (eventId == null || eventId.isBlank()) eventId = UUID.randomUUID().toString();
        if (type == null || type.isBlank()) throw new IllegalArgumentException("Ephemeral event type is required");
        if (unitId == null || unitId.isBlank()) throw new IllegalArgumentException("Ephemeral event unitId is required");
        if (actorId == null || actorId.isBlank()) throw new IllegalArgumentException("Ephemeral event actorId is required");
        if (sessionId == null || sessionId.isBlank()) throw new IllegalArgumentException("Ephemeral event sessionId is required");
        state = Objects.requireNonNullElseGet(state, () -> com.fasterxml.jackson.databind.node.JsonNodeFactory.instance.objectNode());
    }
}
