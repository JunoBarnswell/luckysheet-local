package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.config.AuthProperties;
import com.xc.luckysheet.server.security.LocalAuthSessionRegistry;
import com.xc.luckysheet.server.security.LocalUserAuthentication;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.LocalAuthService;
import com.xc.luckysheet.server.service.ServiceException;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import jakarta.servlet.http.HttpSession;
import jakarta.validation.Valid;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Size;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContext;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.web.context.SecurityContextRepository;
import org.springframework.security.web.csrf.CsrfToken;
import org.springframework.security.web.csrf.CsrfTokenRepository;
import org.springframework.security.authentication.AnonymousAuthenticationToken;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.net.InetAddress;
import java.util.Locale;

@RestController
@RequestMapping("/api/auth")
public class AuthController {
    private final AuthProperties properties;
    private final LocalAuthService auth;
    private final SecurityContextRepository securityContextRepository;
    private final CsrfTokenRepository csrfTokenRepository;
    private final LocalAuthSessionRegistry sessions;

    public AuthController(AuthProperties properties, LocalAuthService auth,
                          SecurityContextRepository securityContextRepository,
                          CsrfTokenRepository csrfTokenRepository,
                          LocalAuthSessionRegistry sessions) {
        this.properties = properties;
        this.auth = auth;
        this.securityContextRepository = securityContextRepository;
        this.csrfTokenRepository = csrfTokenRepository;
        this.sessions = sessions;
    }

    @GetMapping("/config")
    public AuthConfigResponse config() {
        return new AuthConfigResponse(properties.authMode().name().toLowerCase(Locale.ROOT));
    }

    @GetMapping("/session")
    public AuthSessionResponse session(Authentication authentication, HttpServletRequest request,
                                      HttpServletResponse response) {
        boolean authenticated = authentication != null
                && authentication.isAuthenticated()
                && !(authentication instanceof AnonymousAuthenticationToken);
        String subject = null;
        String displayName = null;
        boolean admin = false;
        if (authenticated) {
            subject = ActorIdentity.subject(authentication);
            if (authentication instanceof LocalUserAuthentication local) {
                displayName = local.principalData().displayName();
                admin = local.principalData().admin();
            }
        }
        boolean bootstrapRequired = properties.authMode() == AuthProperties.AuthMode.LOCAL && auth.bootstrapRequired();
        String csrf = properties.authMode() == AuthProperties.AuthMode.LOCAL
                ? csrfToken(request, response).getToken()
                : "";
        return new AuthSessionResponse(authenticated, subject, displayName, admin, bootstrapRequired,
                csrf);
    }

    @PostMapping("/bootstrap")
    public AuthSessionResponse bootstrap(@Valid @RequestBody BootstrapRequest request,
                                         HttpServletRequest servletRequest,
                                         HttpServletResponse servletResponse) {
        requireLoopback(servletRequest);
        LocalUserAuthentication authentication = auth.bootstrap(
                request.token(), request.username(), request.password(), request.displayName());
        saveAuthentication(authentication, servletRequest, servletResponse);
        return session(authentication, servletRequest, servletResponse);
    }

    @PostMapping("/login")
    public AuthSessionResponse login(@Valid @RequestBody LoginRequest request,
                                     HttpServletRequest servletRequest,
                                     HttpServletResponse servletResponse) {
        LocalUserAuthentication authentication = auth.authenticate(request.username(), request.password());
        saveAuthentication(authentication, servletRequest, servletResponse);
        return session(authentication, servletRequest, servletResponse);
    }

    @PostMapping("/logout")
    public ResponseEntity<Void> logout(HttpServletRequest request, HttpServletResponse response,
                                       Authentication authentication) {
        if (authentication instanceof LocalUserAuthentication local) {
            sessions.unregisterHttpSession(local.getName(), request.getSession(false));
        }
        SecurityContext empty = SecurityContextHolder.createEmptyContext();
        SecurityContextHolder.setContext(empty);
        securityContextRepository.saveContext(empty, request, response);
        HttpSession session = request.getSession(false);
        if (session != null) {
            try {
                session.invalidate();
            } catch (IllegalStateException ignored) {
                // The session was already invalidated by the container.
            }
        }
        SecurityContextHolder.clearContext();
        return ResponseEntity.noContent().build();
    }

    private void saveAuthentication(LocalUserAuthentication authentication,
                                    HttpServletRequest request, HttpServletResponse response) {
        HttpSession previousSession = request.getSession(false);
        Authentication previousAuthentication = SecurityContextHolder.getContext().getAuthentication();
        if (previousSession != null) {
            if (previousAuthentication instanceof LocalUserAuthentication previousUser) {
                sessions.unregisterHttpSession(previousUser.getName(), previousSession);
            }
            request.changeSessionId();
        }
        SecurityContext context = SecurityContextHolder.createEmptyContext();
        context.setAuthentication(authentication);
        SecurityContextHolder.setContext(context);
        securityContextRepository.saveContext(context, request, response);
        HttpSession session = request.getSession(false);
        if (session == null) throw new IllegalStateException("Unable to create an authentication session");
        sessions.registerHttpSession(authentication.getName(), session);
    }

    private CsrfToken csrfToken(HttpServletRequest request, HttpServletResponse response) {
        CsrfToken token = (CsrfToken) request.getAttribute(CsrfToken.class.getName());
        if (token != null) return token;
        token = csrfTokenRepository.loadToken(request);
        if (token == null) {
            token = csrfTokenRepository.generateToken(request);
            csrfTokenRepository.saveToken(token, request, response);
        }
        return token;
    }

    private void requireLoopback(HttpServletRequest request) {
        try {
            if (!InetAddress.getByName(request.getRemoteAddr()).isLoopbackAddress()) {
                throw ServiceException.forbidden("Bootstrap is available only from the local machine");
            }
        } catch (java.net.UnknownHostException error) {
            throw ServiceException.forbidden("Bootstrap is available only from the local machine");
        }
    }

    public record AuthConfigResponse(String mode) {
    }

    public record AuthSessionResponse(boolean authenticated, String subject, String displayName,
                                      boolean admin, boolean bootstrapRequired, String csrfToken) {
    }

    public record BootstrapRequest(
            @NotBlank @Size(max = 512) String token,
            @NotBlank @Size(max = 200) String username,
            @NotBlank @Size(max = 200) String password,
            @NotBlank @Size(max = 200) String displayName
    ) {
    }

    public record LoginRequest(
            @NotBlank @Size(max = 200) String username,
            @NotBlank @Size(max = 200) String password
    ) {
    }
}
