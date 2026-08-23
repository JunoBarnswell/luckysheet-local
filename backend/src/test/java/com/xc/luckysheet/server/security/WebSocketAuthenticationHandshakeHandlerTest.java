package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.GuestShareService;
import org.junit.jupiter.api.Test;
import org.springframework.http.HttpHeaders;
import org.springframework.http.server.ServerHttpRequest;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.security.oauth2.jwt.JwtDecoder;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import org.springframework.web.socket.server.HandshakeFailureException;

import java.nio.charset.StandardCharsets;
import java.net.URI;
import java.time.Instant;
import java.util.Base64;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class WebSocketAuthenticationHandshakeHandlerTest {
    @Test
    void bearerSubprotocolIsDecodedAndVerifiedBeforeTheSocketGetsAPrincipal() {
        JwtDecoder decoder = mock(JwtDecoder.class);
        GuestShareService shares = mock(GuestShareService.class);
        Jwt jwt = Jwt.withTokenValue("token").header("alg", "none").subject("owner-1").issuedAt(Instant.now()).expiresAt(Instant.now().plusSeconds(60)).build();
        when(decoder.decode("token")).thenReturn(jwt);
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(decoder, shares);
        String encoded = Base64.getUrlEncoder().withoutPadding().encodeToString("token".getBytes(StandardCharsets.UTF_8));

        var principal = handler.authenticatedPrincipal(request("/ws", "bearer." + encoded));

        assertEquals("owner-1", principal.getName());
        assertInstanceOf(JwtAuthenticationToken.class, principal);
    }

    @Test
    void verifiedShareTokenProducesServerDerivedGuestPrincipal() {
        JwtDecoder decoder = mock(JwtDecoder.class);
        GuestShareService shares = mock(GuestShareService.class);
        GuestShareService.GuestIdentity identity = new GuestShareService.GuestIdentity(
                "guest:share-1", UUID.randomUUID(), "book-1", WorkbookAclRole.COMMENTER, Instant.now().plusSeconds(60)
        );
        when(shares.authenticate("share-token")).thenReturn(identity);
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(decoder, shares);

        var principal = handler.authenticatedPrincipal(request("/ws?shareToken=share-token", null));

        assertEquals("guest:share-1", principal.getName());
        assertInstanceOf(GuestShareAuthentication.class, principal);
    }

    @Test
    void missingCredentialsCannotOpenSocket() {
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(mock(JwtDecoder.class), mock(GuestShareService.class));
        assertThrows(HandshakeFailureException.class, () -> handler.authenticatedPrincipal(request("/ws", null)));
    }

    @Test
    void arbitraryOrAnonymousPrincipalCannotBypassHandshakeAuthentication() {
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(mock(JwtDecoder.class), mock(GuestShareService.class));
        ServerHttpRequest request = request("/ws", null);
        when(request.getPrincipal()).thenReturn(() -> "anonymousUser");

        assertThrows(HandshakeFailureException.class, () -> handler.authenticatedPrincipal(request));
    }

    private ServerHttpRequest request(String path, String protocol) {
        ServerHttpRequest request = mock(ServerHttpRequest.class);
        HttpHeaders headers = new HttpHeaders();
        if (protocol != null) headers.add("Sec-WebSocket-Protocol", protocol);
        when(request.getHeaders()).thenReturn(headers);
        when(request.getURI()).thenReturn(URI.create("http://localhost" + path));
        return request;
    }
}
