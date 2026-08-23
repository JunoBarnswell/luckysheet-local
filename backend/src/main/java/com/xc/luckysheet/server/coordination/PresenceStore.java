package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.config.CoordinationProperties;
import org.springframework.data.redis.core.StringRedisTemplate;

/** Redis is only a short-lived presence/cursor index; it is never workbook state. */
public class PresenceStore {
    private final StringRedisTemplate redis;
    private final ObjectMapper mapper;
    private final CoordinationProperties properties;

    public PresenceStore(StringRedisTemplate redis, ObjectMapper mapper, CoordinationProperties properties) {
        this.redis = redis;
        this.mapper = mapper;
        this.properties = properties;
    }

    public void put(EphemeralEvent event) {
        try {
            String key = key(event);
            redis.opsForValue().set(key, mapper.writeValueAsString(event), properties.presenceTtl());
        } catch (Exception error) {
            throw new IllegalStateException("Unable to persist ephemeral coordination state", error);
        }
    }

    public void remove(EphemeralEvent event) {
        redis.delete(key(event));
    }

    private String key(EphemeralEvent event) {
        String kind = event.type().startsWith("cursor.") ? "cursor" : "presence";
        return "luckysheet:coordination:" + kind + ":" + event.unitId() + ":" + event.actorId() + ":" + event.sessionId();
    }
}
