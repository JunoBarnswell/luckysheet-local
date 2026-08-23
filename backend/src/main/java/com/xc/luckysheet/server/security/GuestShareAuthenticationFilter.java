package com.xc.luckysheet.server.security;

import com.xc.luckysheet.server.service.GuestShareService;
import com.xc.luckysheet.server.service.ServiceException;
import jakarta.servlet.FilterChain;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.web.filter.OncePerRequestFilter;

import java.io.IOException;

/** Authenticates a share token without accepting a client actor or role. */
public class GuestShareAuthenticationFilter extends OncePerRequestFilter {
    public static final String TOKEN_HEADER = "X-Workbook-Share-Token";
    public static final String TOKEN_PARAMETER = "shareToken";

    private final GuestShareService shares;

    public GuestShareAuthenticationFilter(GuestShareService shares) {
        this.shares = shares;
    }

    @Override
    protected void doFilterInternal(HttpServletRequest request, HttpServletResponse response, FilterChain filterChain)
            throws ServletException, IOException {
        // A verified bearer credential remains authoritative. A share token
        // cannot replace or downgrade an authenticated identity.
        if (request.getHeader("Authorization") == null && SecurityContextHolder.getContext().getAuthentication() == null) {
            String token = request.getHeader(TOKEN_HEADER);
            if (token == null || token.isBlank()) {
                token = request.getParameter(TOKEN_PARAMETER);
            }
            if (token != null && !token.isBlank()) {
                try {
                    GuestShareService.GuestIdentity identity = shares.authenticate(token);
                    GuestShareAuthentication authentication = new GuestShareAuthentication(identity);
                    authentication.setDetails(identity);
                    SecurityContextHolder.getContext().setAuthentication(authentication);
                } catch (ServiceException ignored) {
                    // Leave the request unauthenticated; the security chain
                    // returns 401 and does not reveal share validity.
                }
            }
        }
        filterChain.doFilter(request, response);
    }
}
