package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.config.AuthProperties;
import com.xc.luckysheet.server.persistence.LocalUserEntity;
import com.xc.luckysheet.server.persistence.LocalUserEntityRepository;
import com.xc.luckysheet.server.security.LocalAuthSessionRegistry;
import com.xc.luckysheet.server.security.LocalUserAuthentication;
import org.springframework.dao.DataIntegrityViolationException;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.Instant;
import java.util.List;
import java.util.Locale;
import java.util.UUID;

/** Account lifecycle and credential verification for explicit local auth mode. */
@Service
public class LocalAuthService {
    private static final int MAX_USERNAME_LENGTH = 200;
    private static final int MAX_DISPLAY_NAME_LENGTH = 200;
    private static final int MIN_PASSWORD_LENGTH = 12;

    private final AuthProperties properties;
    private final LocalUserEntityRepository users;
    private final PasswordEncoder passwordEncoder;
    private final BootstrapCredentialStore bootstrap;
    private final LocalAuthSessionRegistry sessions;
    private final Object bootstrapLock = new Object();

    public LocalAuthService(AuthProperties properties, LocalUserEntityRepository users,
                            PasswordEncoder passwordEncoder, BootstrapCredentialStore bootstrap,
                            LocalAuthSessionRegistry sessions) {
        this.properties = properties;
        this.users = users;
        this.passwordEncoder = passwordEncoder;
        this.bootstrap = bootstrap;
        this.sessions = sessions;
    }

    public boolean isLocalMode() {
        return properties.authMode() == AuthProperties.AuthMode.LOCAL;
    }

    public boolean bootstrapRequired() {
        requireLocalMode();
        synchronized (bootstrapLock) {
            boolean required = users.count() == 0;
            if (required) bootstrap.ensureToken();
            return required;
        }
    }

    @Transactional
    public LocalUserAuthentication bootstrap(String token, String username, String password, String displayName) {
        requireLocalMode();
        synchronized (bootstrapLock) {
            if (users.count() != 0) throw ServiceException.conflict("Local bootstrap has already completed");
            if (!bootstrap.matches(token)) throw ServiceException.unauthenticated("The bootstrap credential is invalid");
            LocalUserEntity user = saveNewUser(username, password, displayName, true);
            bootstrap.consume();
            return LocalUserAuthentication.from(user);
        }
    }

    public LocalUserAuthentication authenticate(String username, String password) {
        requireLocalMode();
        String normalized = normalizeUsername(username);
        LocalUserEntity user = users.findByUsername(normalized).orElse(null);
        if (user == null || !user.isEnabled() || password == null || !passwordEncoder.matches(password, user.getPasswordHash())) {
            throw ServiceException.unauthenticated("Username or password is invalid");
        }
        return LocalUserAuthentication.from(user);
    }

    @Transactional(readOnly = true)
    public List<LocalUserEntity> listUsers() {
        requireLocalMode();
        return users.findAllByOrderByUsernameAsc();
    }

    @Transactional
    public LocalUserEntity createUser(String username, String password, String displayName, boolean admin) {
        requireLocalMode();
        try {
            return saveNewUser(username, password, displayName, admin);
        } catch (DataIntegrityViolationException error) {
            throw ServiceException.conflict("Username is already in use");
        }
    }

    @Transactional
    public LocalUserEntity setEnabled(String userId, boolean enabled) {
        requireLocalMode();
        LocalUserEntity user = findUser(userId);
        user.setEnabled(enabled);
        LocalUserEntity saved = users.save(user);
        sessions.invalidate(subject(saved));
        return saved;
    }

    @Transactional
    public LocalUserEntity resetPassword(String userId, String password) {
        requireLocalMode();
        LocalUserEntity user = findUser(userId);
        user.resetPassword(passwordEncoder.encode(requirePassword(password)));
        LocalUserEntity saved = users.save(user);
        sessions.invalidate(subject(saved));
        return saved;
    }

    public LocalUserEntity findUser(String userId) {
        if (userId == null || userId.isBlank()) throw ServiceException.validation("User id is required");
        String storageId = userId.startsWith("local:") ? userId.substring("local:".length()) : userId;
        return users.findById(storageId).orElseThrow(() -> ServiceException.notFound("User was not found"));
    }

    public String subject(LocalUserEntity user) {
        return "local:" + user.getUserId();
    }

    public void requireAdmin(LocalUserAuthentication authentication) {
        requireLocalMode();
        if (authentication == null || !authentication.isAuthenticated() || !authentication.principalData().admin()) {
            throw ServiceException.forbidden("Administrator access is required");
        }
    }

    private LocalUserEntity saveNewUser(String username, String password, String displayName, boolean admin) {
        String normalizedUsername = normalizeUsername(username);
        String normalizedDisplayName = requireDisplayName(displayName);
        String encodedPassword = passwordEncoder.encode(requirePassword(password));
        Instant now = Instant.now();
        LocalUserEntity entity = new LocalUserEntity(
                UUID.randomUUID().toString(), normalizedUsername, encodedPassword,
                normalizedDisplayName, true, admin, now, now
        );
        try {
            return users.saveAndFlush(entity);
        } catch (DataIntegrityViolationException error) {
            throw ServiceException.conflict("Username is already in use");
        }
    }

    private void requireLocalMode() {
        if (!isLocalMode()) throw ServiceException.forbidden("Local account authentication is disabled in OIDC mode");
    }

    private String normalizeUsername(String value) {
        if (value == null) throw ServiceException.validation("Username is required");
        String normalized = value.trim().toLowerCase(Locale.ROOT);
        if (normalized.isBlank() || normalized.length() > MAX_USERNAME_LENGTH || normalized.chars().anyMatch(Character::isISOControl)) {
            throw ServiceException.validation("Username is invalid");
        }
        return normalized;
    }

    private String requireDisplayName(String value) {
        if (value == null) throw ServiceException.validation("Display name is required");
        String normalized = value.trim();
        if (normalized.isBlank() || normalized.length() > MAX_DISPLAY_NAME_LENGTH || normalized.chars().anyMatch(Character::isISOControl)) {
            throw ServiceException.validation("Display name is invalid");
        }
        return normalized;
    }

    private String requirePassword(String value) {
        if (value == null || value.length() < MIN_PASSWORD_LENGTH || value.length() > 200
                || value.chars().anyMatch(Character::isISOControl)) {
            throw ServiceException.validation("Password must contain 12 to 200 characters");
        }
        return value;
    }
}
