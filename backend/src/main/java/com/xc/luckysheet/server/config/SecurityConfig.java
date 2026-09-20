package com.xc.luckysheet.server.config;

import com.xc.luckysheet.server.security.ApiSecurityErrorHandlers;
import com.xc.luckysheet.server.security.GuestShareAuthenticationFilter;
import com.xc.luckysheet.server.service.GuestShareService;
import jakarta.servlet.ServletContext;
import org.springframework.boot.web.servlet.ServletContextInitializer;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.config.annotation.method.configuration.EnableMethodSecurity;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.http.SessionCreationPolicy;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.security.oauth2.core.OAuth2Error;
import org.springframework.security.oauth2.core.OAuth2TokenValidator;
import org.springframework.security.oauth2.core.OAuth2TokenValidatorResult;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.security.oauth2.jwt.JwtDecoder;
import org.springframework.security.oauth2.jwt.JwtException;
import org.springframework.security.oauth2.jwt.JwtValidators;
import org.springframework.security.oauth2.jwt.NimbusJwtDecoder;
import org.springframework.security.oauth2.server.resource.web.authentication.BearerTokenAuthenticationFilter;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.web.context.HttpSessionSecurityContextRepository;
import org.springframework.security.web.context.SecurityContextRepository;
import org.springframework.security.web.csrf.CsrfTokenRepository;
import org.springframework.security.web.csrf.HttpSessionCsrfTokenRepository;

@Configuration
@EnableMethodSecurity
@org.springframework.boot.context.properties.EnableConfigurationProperties(AuthProperties.class)
public class SecurityConfig {
    @Bean
    public JwtDecoder jwtDecoder(AuthProperties properties) {
        validateModeConfiguration(properties);
        if (properties.authMode() == AuthProperties.AuthMode.LOCAL) {
            return token -> {
                throw new JwtException("OIDC authentication is disabled in local mode");
            };
        }
        NimbusJwtDecoder decoder = NimbusJwtDecoder.withJwkSetUri(properties.jwksUrl()).build();
        OAuth2TokenValidator<Jwt> issuer = JwtValidators.createDefaultWithIssuer(properties.issuer());
        OAuth2TokenValidator<Jwt> audience = jwt -> jwt.getAudience().stream().anyMatch(properties.audiences()::contains)
                ? OAuth2TokenValidatorResult.success()
                : OAuth2TokenValidatorResult.failure(new OAuth2Error("invalid_token", "Required audience is missing", null));
        decoder.setJwtValidator(new org.springframework.security.oauth2.core.DelegatingOAuth2TokenValidator<>(issuer, audience));
        return decoder;
    }

    @Bean
    public PasswordEncoder passwordEncoder() {
        return new BCryptPasswordEncoder(12);
    }

    @Bean
    public SecurityContextRepository securityContextRepository() {
        return new HttpSessionSecurityContextRepository();
    }

    @Bean
    public CsrfTokenRepository csrfTokenRepository() {
        return new HttpSessionCsrfTokenRepository();
    }

    @Bean
    public ServletContextInitializer sessionCookieSecurityInitializer() {
        return (ServletContext servletContext) -> servletContext.getSessionCookieConfig().setHttpOnly(true);
    }

    @Bean
    public SecurityFilterChain securityFilterChain(
            HttpSecurity http,
            AuthProperties properties,
            GuestShareService shares,
            ApiSecurityErrorHandlers errors,
            SecurityContextRepository securityContextRepository,
            CsrfTokenRepository csrfTokenRepository
    ) throws Exception {
        boolean oidc = properties.authMode() == AuthProperties.AuthMode.OIDC;
        http
                .csrf(csrf -> {
                    if (oidc) csrf.disable();
                    else csrf.csrfTokenRepository(csrfTokenRepository);
                })
                .securityContext(context -> context
                        .securityContextRepository(securityContextRepository)
                        .requireExplicitSave(true))
                .sessionManagement(session -> session.sessionCreationPolicy(
                        oidc ? SessionCreationPolicy.STATELESS : SessionCreationPolicy.IF_REQUIRED))
                .authorizeHttpRequests(auth -> auth
                        .requestMatchers("/health", "/", "/index.html", "/assets/**", "/favicon.ico", "/manifest.webmanifest", "/sw.js").permitAll()
                        .requestMatchers("/api/auth/config", "/api/auth/session", "/api/auth/bootstrap", "/api/auth/login", "/api/auth/logout").permitAll()
                        .requestMatchers("/api/admin/**").hasRole("ADMIN")
                        // WebSocket upgrades authenticate in the dedicated
                        // handshake handler because browser WebSocket APIs
                        // cannot attach a normal Authorization header.
                        .requestMatchers("/ws").permitAll()
                        .requestMatchers("/api/**").authenticated()
                        .anyRequest().permitAll())
                .exceptionHandling(exceptions -> exceptions
                        .authenticationEntryPoint(errors)
                        .accessDeniedHandler(errors));
        http.addFilterBefore(new GuestShareAuthenticationFilter(shares), BearerTokenAuthenticationFilter.class);
        if (oidc) http.oauth2ResourceServer(oauth -> oauth.jwt(jwt -> { }));
        return http.build();
    }

    private void validateModeConfiguration(AuthProperties properties) {
        AuthProperties.AuthMode mode = properties.authMode();
        boolean hasIssuer = properties.issuer() != null && !properties.issuer().isBlank();
        boolean hasAudience = !properties.audiences().isEmpty();
        boolean hasJwks = properties.jwksUrl() != null && !properties.jwksUrl().isBlank();
        if (mode == AuthProperties.AuthMode.LOCAL && (hasIssuer || hasAudience || hasJwks)) {
            throw new IllegalStateException("OIDC settings require LUCKYSHEET_AUTH_MODE=oidc");
        }
        if (mode == AuthProperties.AuthMode.OIDC && (!hasIssuer || !hasAudience || !hasJwks)) {
            throw new IllegalStateException("OIDC mode requires issuer, audience, and jwksUrl");
        }
    }
}
