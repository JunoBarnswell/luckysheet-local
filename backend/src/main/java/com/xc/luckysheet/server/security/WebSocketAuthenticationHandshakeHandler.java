package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.service.GuestShareService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.server.ServerHttpRequest;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.security.oauth2.jwt.JwtDecoder;
import org.springframework.security.oauth2.jwt.JwtException;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.stereotype.Component;
import org.springframework.web.socket.WebSocketHandler;
import org.springframework.web.socket.server.HandshakeFailureException;
import org.springframework.web.socket.server.support.DefaultHandshakeHandler;

import java.nio.charset.StandardCharsets;
import java.security.Principal;
import java.util.Base64;
import java.util.List;
import java.util.Map;

/**
 * Authenticates a browser WebSocket upgrade without accepting a browser actor
 * or role. Browser WebSocket APIs cannot set Authorization headers, so a
 * verified bearer token is carried in the existing base64url subprotocol.
 * Guests use the server-issued credential in a share subprotocol header.
 */
@Component
public final class WebSocketAuthenticationHandshakeHandler extends DefaultHandshakeHandler {
    private static final String BEARER_PROTOCOL_PREFIX = "bearer.";
    private static final String SHARE_PROTOCOL_PREFIX = "share.";
    private static final String WEBSOCKET_PROTOCOL_HEADER = "Sec-WebSocket-Protocol";

    private final JwtDecoder jwtDecoder;
    private final GuestShareService shares;

    public WebSocketAuthenticationHandshakeHandler(JwtDecoder jwtDecoder, GuestShareService shares) {
        this.jwtDecoder = jwtDecoder;
        this.shares = shares;
    }

    @Override
    protected Principal determineUser(ServerHttpRequest request, WebSocketHandler handler, Map<String, Object> attributes) {
        Principal principal = authenticatedPrincipal(request);
        if (principal instanceof LocalUserAuthentication) {
            if (!(request instanceof org.springframework.http.server.ServletServerHttpRequest servlet)
                    || servlet.getServletRequest().getSession(false) == null) throw new HandshakeFailureException("Local authentication session is required");
            attributes.put(LocalAuthSessionRegistry.HTTP_SESSION_ATTRIBUTE, servlet.getServletRequest().getSession(false));
        }
        return principal;
    }

    @Override
    protected String selectProtocol(List<String> requestedProtocols, WebSocketHandler handler) {
        return requestedProtocols.stream()
                .map(String::trim)
                .filter(protocol -> protocol.startsWith(BEARER_PROTOCOL_PREFIX) || protocol.startsWith(SHARE_PROTOCOL_PREFIX))
                .findFirst()
                .orElse(null);
    }

    Principal authenticatedPrincipal(ServerHttpRequest request) {
        Principal existing = request.getPrincipal();
        if (isAuthenticatedPrincipal(existing)) return existing;

        // Servlet WebSocket handshakes normally expose the HTTP session
        // principal through ServerHttpRequest. The security context fallback
        // keeps cookie-authenticated upgrades working with servlet adapters
        // that do not copy request.getUserPrincipal().
        Authentication context = SecurityContextHolder.getContext().getAuthentication();
        if (isAuthenticatedPrincipal(context)) return context;

        String bearer = credentialToken(request.getHeaders(), BEARER_PROTOCOL_PREFIX);
        if (bearer != null) {
            try {
                Jwt jwt = jwtDecoder.decode(bearer);
                if (jwt.getSubject() == null || jwt.getSubject().isBlank()) throw new HandshakeFailureException("Authenticated subject is required");
                return new JwtAuthenticationToken(jwt);
            } catch (JwtException error) {
                throw new HandshakeFailureException("Authentication failed", error);
            }
        }

        String shareToken = credentialToken(request.getHeaders(), SHARE_PROTOCOL_PREFIX);
        if (shareToken != null && !shareToken.isBlank()) {
            try {
                return new GuestShareAuthentication(shares.authenticate(shareToken));
            } catch (RuntimeException error) {
                throw new HandshakeFailureException("Authentication failed", error);
            }
        }
        throw new HandshakeFailureException("Authenticated connection is required");
    }

    private boolean isAuthenticatedPrincipal(Principal principal) {
        return principal instanceof JwtAuthenticationToken
                || principal instanceof GuestShareAuthentication
                || principal instanceof LocalUserAuthentication;
    }

    private String credentialToken(HttpHeaders headers, String prefix) {
        for (String rawHeader : headers.getOrEmpty(WEBSOCKET_PROTOCOL_HEADER)) {
            for (String rawProtocol : rawHeader.split(",")) {
                String protocol = rawProtocol.trim();
                if (!protocol.startsWith(prefix)) continue;
                String encoded = protocol.substring(prefix.length());
                if (encoded.isBlank()) throw new HandshakeFailureException("Authentication failed");
                try {
                    String token = new String(Base64.getUrlDecoder().decode(encoded), StandardCharsets.UTF_8).trim();
                    if (token.isBlank()) throw new HandshakeFailureException("Authentication failed");
                    return token;
                } catch (IllegalArgumentException error) {
                    throw new HandshakeFailureException("Authentication failed", error);
                }
            }
        }
        return null;
    }
}
