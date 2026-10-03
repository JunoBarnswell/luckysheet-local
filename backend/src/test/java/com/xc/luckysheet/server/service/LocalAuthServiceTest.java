package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.config.AuthProperties;
import com.xc.luckysheet.server.persistence.LocalUserEntity;
import com.xc.luckysheet.server.persistence.LocalUserEntityRepository;
import com.xc.luckysheet.server.security.LocalAuthSessionRegistry;
import com.xc.luckysheet.server.security.LocalUserAuthentication;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class LocalAuthServiceTest {
    @TempDir
    Path temp;

    @Test void loginLimitsAccountAliasesAndSourceAttemptsWithoutRejectingNormalAttempts() {
        var properties = new AuthProperties("local", "", "", "", temp.resolve("credential").toString());
        var service = new LocalAuthService(properties, mock(LocalUserEntityRepository.class), new BCryptPasswordEncoder(4), new BootstrapCredentialStore(properties), new LocalAuthSessionRegistry(mock(com.xc.luckysheet.server.service.GuestShareService.class)));
        for (int i = 0; i < 10; i++) assertEquals("UNAUTHENTICATED", assertThrows(ServiceException.class, () -> service.authenticate("Admin", "incorrect-password")).code());
        assertEquals("AUTH_RATE_LIMITED", assertThrows(ServiceException.class, () -> service.authenticate("ADMIN", "incorrect-password")).code());
        for (int i = 0; i < 30; i++) service.reserveSourceAttempt("127.0.0.1");
        assertEquals("AUTH_RATE_LIMITED", assertThrows(ServiceException.class, () -> service.reserveSourceAttempt("127.0.0.1")).code());
        service.reserveSourceAttempt("127.0.0.2");
    }

    @Test void freshBootstrapCredentialIsPrivateBeforeItCanBeRead() throws Exception {
        var path = temp.resolve("fresh");
        new BootstrapCredentialStore(new AuthProperties("local", "", "", "", path.toString())).ensureToken();
        assertTrue(Files.size(path) > 32);
        if (Files.getFileStore(path).supportsFileAttributeView("posix")) assertEquals(java.nio.file.attribute.PosixFilePermissions.fromString("rw-------"), Files.getPosixFilePermissions(path));
    }

    @Test
    void bootstrapHashesPasswordAndConsumesTheOneTimeDiskCredential() throws Exception {
        Path tokenPath = temp.resolve("bootstrap-token");
        Files.writeString(tokenPath, "one-time-token");
        LocalUserEntityRepository users = mock(LocalUserEntityRepository.class);
        when(users.count()).thenReturn(0L);
        when(users.saveAndFlush(any(LocalUserEntity.class))).thenAnswer(invocation -> invocation.getArgument(0));
        var encoder = new BCryptPasswordEncoder(4);
        LocalAuthService service = new LocalAuthService(
                new AuthProperties("local", "", "", "", tokenPath.toString()), users, encoder,
                new BootstrapCredentialStore(new AuthProperties("local", "", "", "", tokenPath.toString())),
                new LocalAuthSessionRegistry(mock(com.xc.luckysheet.server.service.GuestShareService.class))
        );

        LocalUserAuthentication authentication = service.bootstrap(
                "one-time-token", "Admin", "correct horse battery staple", "Administrator");

        assertTrue(authentication.getName().startsWith("local:"));
        assertTrue(authentication.principalData().admin());
        assertFalse(Files.exists(tokenPath));
        var saved = org.mockito.ArgumentCaptor.forClass(LocalUserEntity.class);
        verify(users).saveAndFlush(saved.capture());
        assertTrue(encoder.matches("correct horse battery staple", saved.getValue().getPasswordHash()));
    }

    @Test
    void invalidBootstrapCredentialCannotCreateAnAccount() throws Exception {
        Path tokenPath = temp.resolve("bootstrap-token");
        Files.writeString(tokenPath, "one-time-token");
        LocalUserEntityRepository users = mock(LocalUserEntityRepository.class);
        when(users.count()).thenReturn(0L);
        LocalAuthService service = new LocalAuthService(
                new AuthProperties("local", "", "", "", tokenPath.toString()), users, new BCryptPasswordEncoder(4),
                new BootstrapCredentialStore(new AuthProperties("local", "", "", "", tokenPath.toString())),
                new LocalAuthSessionRegistry(mock(com.xc.luckysheet.server.service.GuestShareService.class))
        );

        assertThrows(ServiceException.class, () -> service.bootstrap("wrong", "admin", "password-123", "Admin"));
        assertTrue(Files.exists(tokenPath));
    }
}
