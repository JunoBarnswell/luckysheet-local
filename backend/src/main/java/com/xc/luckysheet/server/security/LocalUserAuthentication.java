package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.persistence.LocalUserEntity;
import org.springframework.security.authentication.AbstractAuthenticationToken;
import org.springframework.security.core.authority.SimpleGrantedAuthority;

import java.util.List;

/** Authenticated principal backed by a persisted local account. */
public final class LocalUserAuthentication extends AbstractAuthenticationToken {
    private final PrincipalData principal;

    private LocalUserAuthentication(PrincipalData principal) {
        super(principal.admin() ? List.of(new SimpleGrantedAuthority("ROLE_ADMIN")) : List.of());
        this.principal = principal;
        setAuthenticated(true);
    }

    public static LocalUserAuthentication from(LocalUserEntity user) {
        return new LocalUserAuthentication(new PrincipalData(
                "local:" + user.getUserId(),
                user.getUsername(),
                user.getDisplayName(),
                user.isAdmin(),
                user.getCredentialVersion()
        ));
    }

    @Override
    public Object getCredentials() {
        return null;
    }

    @Override
    public Object getPrincipal() {
        return principal;
    }

    @Override
    public String getName() {
        return principal.subject();
    }

    public PrincipalData principalData() {
        return principal;
    }

    public record PrincipalData(String subject, String username, String displayName,
                                boolean admin, long credentialVersion) {
    }
}
