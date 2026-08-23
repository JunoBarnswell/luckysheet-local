package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.JsonNode;
import org.springframework.beans.factory.ObjectProvider;
import org.springframework.stereotype.Component;

import java.util.UUID;

/** Publishes transient state when Redis is configured and remains local otherwise. */
@Component
public class EphemeralCoordinationService {
    private final ObjectProvider<RedisCoordinationPublisher> publisher;
    private final ObjectProvider<PresenceStore> presence;

    public EphemeralCoordinationService(
            ObjectProvider<RedisCoordinationPublisher> publisher,
            ObjectProvider<PresenceStore> presence
    ) {
        this.publisher = publisher;
        this.presence = presence;
    }

    public EphemeralEvent updated(String type, String unitId, String actorId, String sessionId, JsonNode state) {
        EphemeralEvent event = new EphemeralEvent(UUID.randomUUID().toString(), type, unitId, actorId, sessionId, state);
        presence.ifAvailable(store -> store.put(event));
        publisher.ifAvailable(redis -> redis.publishEphemeral(event));
        return event;
    }

    public EphemeralEvent offline(String unitId, String actorId, String sessionId) {
        EphemeralEvent event = new EphemeralEvent(
                UUID.randomUUID().toString(), "presence.updated", unitId, actorId, sessionId,
                com.fasterxml.jackson.databind.node.JsonNodeFactory.instance.objectNode().put("status", "offline")
        );
        presence.ifAvailable(store -> store.remove(event));
        publisher.ifAvailable(redis -> redis.publishEphemeral(event));
        return event;
    }
}
