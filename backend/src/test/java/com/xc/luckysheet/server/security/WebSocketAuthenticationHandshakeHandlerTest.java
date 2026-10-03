package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.contract.WorkbookRole;
import com.xc.luckysheet.server.persistence.LocalUserEntity;
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
        Jwt jwt = Jwt.withTokenValue("token").header("alg", "none").subject("owner-1").issuer("https://issuer.test").claim("sid", "sso-session").issuedAt(Instant.now()).expiresAt(Instant.now().plusSeconds(60)).build();
        when(decoder.decode("token")).thenReturn(jwt);
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(decoder, shares, identities());
        String encoded = Base64.getUrlEncoder().withoutPadding().encodeToString("token".getBytes(StandardCharsets.UTF_8));

        var principal = handler.authenticatedPrincipal(request("/ws", "bearer." + encoded));

        assertEquals(identities().jwtContext(jwt).principal(), principal.getName());
        assertInstanceOf(JwtAuthenticationToken.class, principal);
    }

    @Test
    void verifiedShareTokenProducesServerDerivedGuestPrincipal() {
        JwtDecoder decoder = mock(JwtDecoder.class);
        GuestShareService shares = mock(GuestShareService.class);
        GuestShareService.GuestIdentity identity = new GuestShareService.GuestIdentity(
                "guest:share-1", UUID.randomUUID(), "book-1", WorkbookRole.COMMENTER, Instant.now().plusSeconds(60)
        );
        when(shares.authenticate("share-token")).thenReturn(identity);
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(decoder, shares, identities());

        var principal = handler.authenticatedPrincipal(request("/ws", "share." + Base64.getUrlEncoder().withoutPadding().encodeToString("share-token".getBytes(StandardCharsets.UTF_8))));

        assertEquals("guest:share-1", principal.getName());
        assertInstanceOf(GuestShareAuthentication.class, principal);
    }

    @Test
    void queryCredentialsAreNotAccepted() {
        var handler = new WebSocketAuthenticationHandshakeHandler(mock(JwtDecoder.class), mock(GuestShareService.class));
        assertThrows(HandshakeFailureException.class, () -> handler.authenticatedPrincipal(request("/ws?shareToken=share-token", null)));
    }

    @Test
    void missingCredentialsCannotOpenSocket() {
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(mock(JwtDecoder.class), mock(GuestShareService.class), identities());
        assertThrows(HandshakeFailureException.class, () -> handler.authenticatedPrincipal(request("/ws", null)));
    }

    @Test
    void arbitraryOrAnonymousPrincipalCannotBypassHandshakeAuthentication() {
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(mock(JwtDecoder.class), mock(GuestShareService.class), identities());
        ServerHttpRequest request = request("/ws", null);
        when(request.getPrincipal()).thenReturn(() -> "anonymousUser");

        assertThrows(HandshakeFailureException.class, () -> handler.authenticatedPrincipal(request));
    }

    @Test
    void sessionCookiePrincipalIsPassedToTheSocketWithItsStableLocalSubject() {
        JwtDecoder decoder = mock(JwtDecoder.class);
        GuestShareService shares = mock(GuestShareService.class);
        LocalUserEntity user = new LocalUserEntity(
                "user-1", "owner", "hash", "Owner", true, true,
                Instant.now(), Instant.now()
        );
        WebSocketAuthenticationHandshakeHandler handler = new WebSocketAuthenticationHandshakeHandler(decoder, shares, identities());
        ServerHttpRequest request = request("/ws", null);
        when(request.getPrincipal()).thenReturn(LocalUserAuthentication.from(user));

        var principal = handler.authenticatedPrincipal(request);

        assertEquals("local:user-1", principal.getName());
        assertInstanceOf(LocalUserAuthentication.class, principal);
    }

    private VerifiedIdentityService identities() {
        return new VerifiedIdentityService(new com.xc.luckysheet.server.config.IdentityProperties(null, null, null, null, null, null, null));
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
