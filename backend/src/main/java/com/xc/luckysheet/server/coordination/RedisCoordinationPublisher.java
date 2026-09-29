package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.xc.luckysheet.server.config.CoordinationProperties;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.xc.luckysheet.server.store.OutboxRow;
import org.springframework.data.redis.core.StringRedisTemplate;

/** Publishes coordination messages; PostgreSQL remains the source of truth. */
public class RedisCoordinationPublisher {
    private final StringRedisTemplate redis;
    private final ObjectMapper mapper;
    private final CoordinationProperties properties;

    public RedisCoordinationPublisher(StringRedisTemplate redis, ObjectMapper mapper, CoordinationProperties properties) {
        this.redis = redis;
        this.mapper = mapper;
        this.properties = properties;
    }

    public void publishRevision(OutboxRow event) {
        try {
            CommittedOperationEnvelope operation = mapper.readValue(event.payloadJson(), CommittedOperationEnvelope.class);
            var message = new RevisionCoordinationEvent("revision", event.eventId().toString(), event.unitId(),
                    event.operationId(), event.revision(), operation);
            redis.convertAndSend(properties.channel(), mapper.writeValueAsString(message));
        } catch (Exception error) {
            throw new IllegalStateException("Unable to publish revision coordination event", error);
        }
    }

    public void publishEphemeral(EphemeralEvent event) {
        try {
            ObjectNode message = mapper.createObjectNode()
                    .put("kind", "ephemeral")
                    .put("eventId", event.eventId())
                    .put("type", event.type())
                    .put("unitId", event.unitId())
                    .put("actorId", event.actorId())
                    .put("sessionId", event.sessionId());
            message.set("state", event.state().deepCopy());
            redis.convertAndSend(properties.channel(), mapper.writeValueAsString(message));
        } catch (Exception error) {
            throw new IllegalStateException("Unable to publish ephemeral coordination event", error);
        }
    }
}
