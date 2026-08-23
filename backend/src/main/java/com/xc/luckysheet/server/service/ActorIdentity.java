package com.xc.luckysheet.server.service;

import org.springframework.security.core.Authentication;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import com.xc.luckysheet.server.security.GuestShareAuthentication;

import java.security.Principal;

public final class ActorIdentity {
    private ActorIdentity() {
    }

    public static String subject(Authentication authentication) {
        if (authentication instanceof JwtAuthenticationToken token && token.getToken().getSubject() != null && !token.getToken().getSubject().isBlank()) {
            return token.getToken().getSubject();
        }
        if (authentication instanceof GuestShareAuthentication guest) return guest.getName();
        throw ServiceException.forbidden("Authenticated subject is required");
    }

    public static void requireRegisteredActor(Authentication authentication) {
        if (authentication instanceof GuestShareAuthentication) {
            throw ServiceException.forbidden("Guest shares cannot create workbooks or manage ACL");
        }
    }

    public static String subject(Principal principal) {
        if (principal != null && principal.getName() != null && !principal.getName().isBlank()) return principal.getName();
        throw ServiceException.forbidden("Authenticated subject is required");
    }
}
