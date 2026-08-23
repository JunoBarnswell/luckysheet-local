package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.service.GuestShareService;
import org.springframework.security.authentication.AbstractAuthenticationToken;
import org.springframework.security.core.authority.SimpleGrantedAuthority;

import java.util.List;

public final class GuestShareAuthentication extends AbstractAuthenticationToken {
    private final GuestShareService.GuestIdentity identity;

    public GuestShareAuthentication(GuestShareService.GuestIdentity identity) {
        super(List.of(new SimpleGrantedAuthority("ROLE_GUEST")));
        this.identity = identity;
        setAuthenticated(true);
    }

    @Override
    public Object getCredentials() {
        return null;
    }

    @Override
    public Object getPrincipal() {
        return identity;
    }

    @Override
    public String getName() {
        return identity.subject();
    }

    public GuestShareService.GuestIdentity identity() {
        return identity;
    }
}
