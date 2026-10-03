package com.xc.luckysheet.server.service;

import org.springframework.security.core.Authentication;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import com.xc.luckysheet.server.security.GuestShareAuthentication;
import com.xc.luckysheet.server.security.LocalUserAuthentication;

import java.security.Principal;
import java.util.Collection;
import java.util.List;

public final class ActorIdentity {
    private ActorIdentity() {
    }

    public static String subject(Authentication authentication) {
        if (authentication instanceof JwtAuthenticationToken token && token.getName() != null && !token.getName().isBlank()) {
            return token.getName();
        }
        if (authentication instanceof LocalUserAuthentication local && local.getName() != null && !local.getName().isBlank()) {
            return local.getName();
        }
        if (authentication instanceof GuestShareAuthentication guest) return guest.getName();
        throw ServiceException.forbidden("Authenticated subject is required");
    }

    public static void requireRegisteredActor(Authentication authentication) {
        if (authentication instanceof GuestShareAuthentication) {
            throw ServiceException.forbidden("Guest shares cannot create workbooks or manage ACL");
        }
    }

    /** Group identifiers are accepted only from the verified JWT claim set. */
    public static List<String> groups(Authentication authentication) {
        if (!(authentication instanceof JwtAuthenticationToken token)) return List.of();
        Object raw = token.getToken().getClaims().get("groups");
        if (!(raw instanceof Collection<?> values)) return List.of();
        return values.stream().filter(String.class::isInstance).map(String.class::cast)
                .map(String::trim).filter(value -> !value.isEmpty()).distinct().toList();
    }

    public static List<String> groups(Principal principal) {
        return principal instanceof Authentication authentication ? groups(authentication) : List.of();
    }

    public static String subject(Principal principal) {
        if (principal instanceof Authentication authentication) return subject(authentication);
        if (principal != null && principal.getName() != null && !principal.getName().isBlank()) return principal.getName();
        throw ServiceException.forbidden("Authenticated subject is required");
    }
}
