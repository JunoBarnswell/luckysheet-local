package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.persistence.LocalUserEntity;
import com.xc.luckysheet.server.security.LocalUserAuthentication;
import com.xc.luckysheet.server.service.LocalAuthService;
import jakarta.validation.Valid;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotNull;
import jakarta.validation.constraints.Size;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PatchMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;

@RestController
@RequestMapping("/api/admin/users")
public class AdminUserController {
    private final LocalAuthService auth;

    public AdminUserController(LocalAuthService auth) {
        this.auth = auth;
    }

    @GetMapping
    public List<UserResponse> list(Authentication authentication) {
        auth.requireAdmin(local(authentication));
        return auth.listUsers().stream().map(UserResponse::from).toList();
    }

    @PostMapping
    public UserResponse create(@Valid @RequestBody CreateUserRequest request, Authentication authentication) {
        auth.requireAdmin(local(authentication));
        return UserResponse.from(auth.createUser(request.username(), request.password(), request.displayName(), request.admin()));
    }

    @PatchMapping("/{id}")
    public UserResponse setEnabled(@PathVariable String id, @Valid @RequestBody EnabledRequest request,
                                   Authentication authentication) {
        auth.requireAdmin(local(authentication));
        return UserResponse.from(auth.setEnabled(id, request.enabled()));
    }

    @PostMapping("/{id}/password")
    public UserResponse resetPassword(@PathVariable String id, @Valid @RequestBody PasswordRequest request,
                                      Authentication authentication) {
        auth.requireAdmin(local(authentication));
        return UserResponse.from(auth.resetPassword(id, request.password()));
    }

    private LocalUserAuthentication local(Authentication authentication) {
        return authentication instanceof LocalUserAuthentication value ? value : null;
    }

    public record CreateUserRequest(
            @NotBlank @Size(max = 200) String username,
            @NotBlank @Size(max = 200) String password,
            @NotBlank @Size(max = 200) String displayName,
            boolean admin
    ) {
    }

    public record EnabledRequest(@NotNull Boolean enabled) {
    }

    public record PasswordRequest(@NotBlank @Size(max = 200) String password) {
    }

    public record UserResponse(String id, String username, String displayName, boolean enabled, boolean admin) {
        static UserResponse from(LocalUserEntity user) {
            return new UserResponse("local:" + user.getUserId(), user.getUsername(), user.getDisplayName(), user.isEnabled(), user.isAdmin());
        }
    }
}
