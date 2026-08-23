package com.xc.luckysheet.server.coordination;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.data.redis.connection.Message;
import org.springframework.data.redis.connection.MessageListener;

/** Converts Redis notifications into delivery to this instance's WebSockets. */
public class RedisCoordinationSubscriber implements MessageListener {
    private static final Logger LOGGER = LoggerFactory.getLogger(RedisCoordinationSubscriber.class);

    private final ObjectMapper mapper;
    private final WebSocketSessionRegistry sessions;

    public RedisCoordinationSubscriber(ObjectMapper mapper, WebSocketSessionRegistry sessions) {
        this.mapper = mapper;
        this.sessions = sessions;
    }

    @Override
    public void onMessage(Message message, byte[] pattern) {
        try {
            JsonNode root = mapper.readTree(message.getBody());
            switch (root.path("kind").asText()) {
                case "revision" -> {
                    CommittedOperationEnvelope operation = mapper.treeToValue(root.path("operation"), CommittedOperationEnvelope.class);
                    sessions.broadcastRevision(operation);
                }
                case "ephemeral" -> {
                    EphemeralEvent event = new EphemeralEvent(
                            root.path("eventId").asText(),
                            root.path("type").asText(),
                            root.path("unitId").asText(),
                            root.path("actorId").asText(),
                            root.path("sessionId").asText(),
                            root.path("state")
                    );
                    sessions.broadcastEphemeral(event);
                }
                default -> LOGGER.debug("Ignoring unknown coordination event kind");
            }
        } catch (Exception error) {
            // A malformed notification cannot become workbook state or break
            // the Redis listener thread; the durable PostgreSQL record remains
            // available for recovery and diagnostics.
            LOGGER.warn("Ignoring malformed coordination notification", error);
        }
    }
}
