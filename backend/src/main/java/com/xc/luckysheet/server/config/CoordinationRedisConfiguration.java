package com.xc.luckysheet.server.config;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.coordination.PresenceStore;
import com.xc.luckysheet.server.coordination.RedisCoordinationPublisher;
import com.xc.luckysheet.server.coordination.RedisCoordinationSubscriber;
import com.xc.luckysheet.server.coordination.WebSocketSessionRegistry;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Primary;
import org.springframework.data.redis.connection.RedisConnectionFactory;
import org.springframework.data.redis.listener.ChannelTopic;
import org.springframework.data.redis.listener.RedisMessageListenerContainer;
import org.springframework.data.redis.core.StringRedisTemplate;

@Configuration
@ConditionalOnProperty(prefix = "luckysheet.coordination", name = "redis-enabled", havingValue = "true")
public class CoordinationRedisConfiguration {
    @Bean
    @Primary
    public StringRedisTemplate coordinationRedisTemplate(RedisConnectionFactory connectionFactory) {
        return new StringRedisTemplate(connectionFactory);
    }

    @Bean
    public RedisCoordinationPublisher redisCoordinationPublisher(
            StringRedisTemplate coordinationRedisTemplate,
            ObjectMapper mapper,
            CoordinationProperties properties
    ) {
        return new RedisCoordinationPublisher(coordinationRedisTemplate, mapper, properties);
    }

    @Bean
    public PresenceStore presenceStore(
            StringRedisTemplate coordinationRedisTemplate,
            ObjectMapper mapper,
            CoordinationProperties properties
    ) {
        return new PresenceStore(coordinationRedisTemplate, mapper, properties);
    }

    @Bean
    public RedisCoordinationSubscriber redisCoordinationSubscriber(ObjectMapper mapper, WebSocketSessionRegistry sessions) {
        return new RedisCoordinationSubscriber(mapper, sessions);
    }

    @Bean(initMethod = "start", destroyMethod = "stop")
    public RedisMessageListenerContainer coordinationRedisListenerContainer(
            RedisConnectionFactory connectionFactory,
            CoordinationProperties properties,
            RedisCoordinationSubscriber subscriber
    ) {
        RedisMessageListenerContainer container = new RedisMessageListenerContainer();
        container.setConnectionFactory(connectionFactory);
        container.addMessageListener(subscriber, new ChannelTopic(properties.channel()));
        return container;
    }
}
