package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.service.GuestShareService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.server.ServerHttpRequest;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.security.oauth2.jwt.JwtDecoder;
import org.springframework.security.oauth2.jwt.JwtException;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import org.springframework.stereotype.Component;
import org.springframework.util.MultiValueMap;
import org.springframework.web.socket.WebSocketHandler;
import org.springframework.web.socket.server.HandshakeFailureException;
import org.springframework.web.socket.server.support.DefaultHandshakeHandler;
import org.springframework.web.util.UriComponentsBuilder;

import java.nio.charset.StandardCharsets;
import java.security.Principal;
import java.util.Base64;
import java.util.Map;

/**
 * Authenticates a browser WebSocket upgrade without accepting a browser actor
 * or role. Browser WebSocket APIs cannot set Authorization headers, so a
 * verified bearer token is carried in the existing base64url subprotocol.
 * Share links use only the server-issued share token query parameter.
 */
@Component
public final class WebSocketAuthenticationHandshakeHandler extends DefaultHandshakeHandler {
    private static final String BEARER_PROTOCOL_PREFIX = "bearer.";
    private static final String WEBSOCKET_PROTOCOL_HEADER = "Sec-WebSocket-Protocol";

    private final JwtDecoder jwtDecoder;
    private final GuestShareService shares;

    public WebSocketAuthenticationHandshakeHandler(JwtDecoder jwtDecoder, GuestShareService shares) {
        this.jwtDecoder = jwtDecoder;
        this.shares = shares;
    }

    @Override
    protected Principal determineUser(ServerHttpRequest request, WebSocketHandler handler, Map<String, Object> attributes) {
        return authenticatedPrincipal(request);
    }

    Principal authenticatedPrincipal(ServerHttpRequest request) {
        Principal existing = request.getPrincipal();
        if (existing instanceof JwtAuthenticationToken || existing instanceof GuestShareAuthentication) return existing;

        String bearer = bearerToken(request.getHeaders());
        if (bearer != null) {
            try {
                Jwt jwt = jwtDecoder.decode(bearer);
                if (jwt.getSubject() == null || jwt.getSubject().isBlank()) throw new HandshakeFailureException("Authenticated subject is required");
                return new JwtAuthenticationToken(jwt);
            } catch (JwtException error) {
                throw new HandshakeFailureException("Authentication failed", error);
            }
        }

        MultiValueMap<String, String> query = UriComponentsBuilder.fromUri(request.getURI()).build().getQueryParams();
        String shareToken = query.getFirst(GuestShareAuthenticationFilter.TOKEN_PARAMETER);
        if (shareToken != null && !shareToken.isBlank()) {
            try {
                return new GuestShareAuthentication(shares.authenticate(shareToken));
            } catch (RuntimeException error) {
                throw new HandshakeFailureException("Authentication failed", error);
            }
        }
        throw new HandshakeFailureException("Authenticated connection is required");
    }

    private String bearerToken(HttpHeaders headers) {
        for (String rawHeader : headers.getOrEmpty(WEBSOCKET_PROTOCOL_HEADER)) {
            for (String rawProtocol : rawHeader.split(",")) {
                String protocol = rawProtocol.trim();
                if (!protocol.startsWith(BEARER_PROTOCOL_PREFIX)) continue;
                String encoded = protocol.substring(BEARER_PROTOCOL_PREFIX.length());
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
