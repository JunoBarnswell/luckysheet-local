package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Configuration;

@Configuration
@EnableConfigurationProperties({CoordinationProperties.class, QueryProperties.class, ShareProperties.class})
public class CoordinationConfig {
}
