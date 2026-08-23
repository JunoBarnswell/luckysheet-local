package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.service.GuestShareService;
import jakarta.servlet.FilterChain;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockHttpServletRequest;
import org.springframework.mock.web.MockHttpServletResponse;
import org.springframework.mock.web.MockFilterChain;
import org.springframework.security.core.context.SecurityContextHolder;

import java.time.Instant;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class GuestShareAuthenticationFilterTest {
    @AfterEach
    void clearSecurityContext() {
        SecurityContextHolder.clearContext();
    }

    @Test
    void validShareTokenCreatesServerDerivedGuestIdentity() throws Exception {
        GuestShareService shares = mock(GuestShareService.class);
        GuestShareService.GuestIdentity identity = new GuestShareService.GuestIdentity(
                "guest:share-1", UUID.randomUUID(), "unit-1", WorkbookAclRole.VIEWER, Instant.now().plusSeconds(60)
        );
        when(shares.authenticate("token")).thenReturn(identity);
        GuestShareAuthenticationFilter filter = new GuestShareAuthenticationFilter(shares);
        MockHttpServletRequest request = new MockHttpServletRequest("GET", "/api/workbooks/unit-1/snapshot");
        request.addHeader(GuestShareAuthenticationFilter.TOKEN_HEADER, "token");
        filter.doFilter(request, new MockHttpServletResponse(), new MockFilterChain());

        GuestShareAuthentication authentication = assertInstanceOf(GuestShareAuthentication.class, SecurityContextHolder.getContext().getAuthentication());
        assertEquals("guest:share-1", authentication.getName());
        assertEquals("ROLE_GUEST", authentication.getAuthorities().iterator().next().getAuthority());
    }
}
