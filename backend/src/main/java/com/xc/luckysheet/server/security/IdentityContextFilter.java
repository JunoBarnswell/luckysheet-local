package com.xc.luckysheet.server.security;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.ApiErrorResponse;
import jakarta.servlet.FilterChain;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.web.filter.OncePerRequestFilter;

import java.io.IOException;

/** Refuses an SDK request before a mutation can execute under a different cookie/token context. */
public final class IdentityContextFilter extends OncePerRequestFilter {
    private final VerifiedIdentityService identities;
    private final ObjectMapper mapper;
    public IdentityContextFilter(VerifiedIdentityService identities, ObjectMapper mapper) {
        this.identities = identities; this.mapper = mapper;
    }
    @Override
    protected void doFilterInternal(HttpServletRequest request, HttpServletResponse response, FilterChain chain)
            throws ServletException, IOException {
        String expected = request.getHeader("X-Spreadsheet-Context");
        if (expected != null) {
            Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
            try {
                if (!expected.equals(identities.context(authentication, request).contextId())) throw new IllegalStateException("Identity context changed");
            } catch (RuntimeException error) {
                response.setStatus(409); response.setContentType("application/json");
                mapper.writeValue(response.getOutputStream(), new ApiErrorResponse("STALE_OPERATION", "Verified identity context changed; reopen the workbook"));
                return;
            }
        }
        chain.doFilter(request, response);
    }
}
