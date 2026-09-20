package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.config.AuthProperties;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.util.Base64;

/** Owns the one-time local bootstrap credential stored outside the database. */
@Component
public final class BootstrapCredentialStore {
    private final Path path;
    private final SecureRandom random = new SecureRandom();
    private final Object lock = new Object();

    public BootstrapCredentialStore(AuthProperties properties) {
        this.path = properties.bootstrapPath();
    }

    public void ensureToken() {
        synchronized (lock) {
            if (Files.exists(path)) return;
            try {
                Path parent = path.getParent();
                if (parent != null) Files.createDirectories(parent);
                String token = Base64.getUrlEncoder().withoutPadding().encodeToString(random.generateSeed(32));
                Files.writeString(path, token, StandardCharsets.UTF_8, StandardOpenOption.CREATE_NEW, StandardOpenOption.WRITE);
            } catch (java.nio.file.FileAlreadyExistsException ignored) {
                // Another request created the credential while this request raced.
            } catch (IOException error) {
                throw new IllegalStateException("Unable to create the local bootstrap credential", error);
            }
        }
    }

    public boolean matches(String candidate) {
        if (candidate == null || candidate.isBlank()) return false;
        synchronized (lock) {
            try {
                if (!Files.isRegularFile(path)) return false;
                byte[] expected = Files.readAllBytes(path);
                byte[] actual = candidate.getBytes(StandardCharsets.UTF_8);
                return MessageDigest.isEqual(expected, actual);
            } catch (IOException error) {
                throw new IllegalStateException("Unable to read the local bootstrap credential", error);
            }
        }
    }

    public void consume() {
        synchronized (lock) {
            try {
                Files.deleteIfExists(path);
            } catch (IOException error) {
                throw new IllegalStateException("Unable to consume the local bootstrap credential", error);
            }
        }
    }
}
