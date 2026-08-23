package com.xc.luckysheet.server.config;

import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.config.annotation.method.configuration.EnableMethodSecurity;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.http.SessionCreationPolicy;
import org.springframework.security.oauth2.core.OAuth2Error;
import org.springframework.security.oauth2.core.OAuth2TokenValidator;
import org.springframework.security.oauth2.core.OAuth2TokenValidatorResult;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.security.oauth2.jwt.JwtDecoder;
import org.springframework.security.oauth2.jwt.JwtValidators;
import org.springframework.security.oauth2.jwt.NimbusJwtDecoder;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.oauth2.server.resource.web.authentication.BearerTokenAuthenticationFilter;
import com.xc.luckysheet.server.security.GuestShareAuthenticationFilter;
import com.xc.luckysheet.server.service.GuestShareService;
import com.xc.luckysheet.server.security.ApiSecurityErrorHandlers;

@Configuration
@EnableMethodSecurity
@EnableConfigurationProperties(AuthProperties.class)
public class SecurityConfig {
    @Bean
    public JwtDecoder jwtDecoder(AuthProperties properties) {
        if (properties.issuer() == null || properties.issuer().isBlank()) throw new IllegalStateException("AUTH_ISSUER must be configured");
        if (properties.jwksUrl() == null || properties.jwksUrl().isBlank()) throw new IllegalStateException("AUTH_JWKS_URL must be configured");
        if (properties.audiences().isEmpty()) throw new IllegalStateException("AUTH_AUDIENCE must be configured");
        NimbusJwtDecoder decoder = NimbusJwtDecoder.withJwkSetUri(properties.jwksUrl()).build();
        OAuth2TokenValidator<Jwt> issuer = JwtValidators.createDefaultWithIssuer(properties.issuer());
        OAuth2TokenValidator<Jwt> audience = jwt -> jwt.getAudience().stream().anyMatch(properties.audiences()::contains)
                ? OAuth2TokenValidatorResult.success()
                : OAuth2TokenValidatorResult.failure(new OAuth2Error("invalid_token", "Required audience is missing", null));
        decoder.setJwtValidator(new org.springframework.security.oauth2.core.DelegatingOAuth2TokenValidator<>(issuer, audience));
        return decoder;
    }

    @Bean
    public SecurityFilterChain securityFilterChain(HttpSecurity http, GuestShareService shares, ApiSecurityErrorHandlers errors) throws Exception {
        http
                .csrf(csrf -> csrf.disable())
                .sessionManagement(session -> session.sessionCreationPolicy(SessionCreationPolicy.STATELESS))
                .authorizeHttpRequests(auth -> auth
                        .requestMatchers("/health").permitAll()
                        .requestMatchers("/api/**").authenticated()
                        // WebSocket upgrades authenticate in the dedicated
                        // handshake handler because browser WebSocket APIs
                        // cannot attach a normal Authorization header.
                        .requestMatchers("/ws").permitAll()
                        .anyRequest().denyAll())
                .oauth2ResourceServer(oauth -> oauth.jwt(jwt -> { }));
        http.exceptionHandling(exceptions -> exceptions
                .authenticationEntryPoint(errors)
                .accessDeniedHandler(errors));
        http.addFilterBefore(new GuestShareAuthenticationFilter(shares), BearerTokenAuthenticationFilter.class);
        return http.build();
    }
}
