package com.xc.luckysheet.server.service;

import com.xc.luckysheet.server.config.ShareProperties;
import com.xc.luckysheet.server.contract.ShareCreateRequest;
import com.xc.luckysheet.server.contract.ShareResponse;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.ShareRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.security.SecureRandom;
import java.time.Instant;
import java.util.Base64;
import java.util.List;
import java.util.UUID;

@Service
public class GuestShareService {
    private final WorkbookStore store;
    private final ShareProperties properties;
    private final WorkbookLifecycleService lifecycle;
    private final WorkbookAuthorizationService authorization;
    private final SecureRandom random = new SecureRandom();

    public GuestShareService(WorkbookStore store, ShareProperties properties, WorkbookLifecycleService lifecycle,
                             WorkbookAuthorizationService authorization) {
        this.store = store;
        this.properties = properties;
        this.lifecycle = lifecycle;
        this.authorization = authorization;
    }

    @Transactional
    public ShareResponse create(String unitId, ShareCreateRequest request, String actor) {
        // The owner check must use the persistent ACL, never a client role.
        requireOwner(unitId, actor);
        lifecycle.requireActive(unitId);
        WorkbookAclRole role = parseRole(request.role());
        Instant now = Instant.now();
        Instant expiresAt = request.expiresAt() == null ? now.plus(properties.defaultLifetime()) : request.expiresAt();
        if (!expiresAt.isAfter(now)) throw ServiceException.validation("Share expiry must be in the future");
        if (expiresAt.isAfter(now.plus(properties.maxLifetime()))) throw ServiceException.validation("Share expiry exceeds the configured maximum lifetime");
        UUID shareId = UUID.randomUUID();
        String secret = randomSecret();
        String token = shareId + "." + secret;
        ShareRow row = new ShareRow(shareId, unitId, hash(token), role, expiresAt, null, actor, now);
        store.insertShare(row);
        return new ShareResponse(shareId, unitId, role, expiresAt, null, actor, now, token);
    }

    public List<ShareResponse> list(String unitId, String actor) {
        requireOwner(unitId, actor);
        lifecycle.requireActive(unitId);
        return store.listShares(unitId).stream().map(ShareResponse::listed).toList();
    }

    @Transactional
    public void revoke(String unitId, UUID shareId, String actor) {
        requireOwner(unitId, actor);
        lifecycle.requireActive(unitId);
        if (store.revokeShare(unitId, shareId, Instant.now()) == 0) throw ServiceException.notFound("Share link not found or already revoked");
    }

    public GuestIdentity authenticate(String token) {
        if (token == null || token.isBlank()) throw ServiceException.unauthenticated("A share token is required");
        int separator = token.indexOf('.');
        if (separator <= 0 || separator == token.length() - 1) throw ServiceException.unauthenticated("Share token is invalid");
        UUID shareId;
        try {
            shareId = UUID.fromString(token.substring(0, separator));
        } catch (IllegalArgumentException error) {
            throw ServiceException.unauthenticated("Share token is invalid");
        }
        ShareRow row = store.findActiveShare(shareId, Instant.now()).orElseThrow(() -> ServiceException.unauthenticated("Share token is expired or revoked"));
        byte[] expected = hash(token).getBytes(StandardCharsets.US_ASCII);
        byte[] actual = row.tokenHash().getBytes(StandardCharsets.US_ASCII);
        // The database stores only the digest; compare its textual form in constant time.
        if (!MessageDigest.isEqual(expected, actual)) throw ServiceException.unauthenticated("Share token is invalid");
        return new GuestIdentity("guest:" + row.shareId(), row.shareId(), row.unitId(), row.role(), row.expiresAt());
    }

    public WorkbookAclRole roleFor(String unitId, String subject) {
        if (subject == null || !subject.startsWith("guest:")) return null;
        UUID shareId;
        try {
            shareId = UUID.fromString(subject.substring("guest:".length()));
        } catch (IllegalArgumentException error) {
            return null;
        }
        return store.findActiveShare(unitId, shareId, Instant.now()).map(ShareRow::role).orElse(null);
    }

    private WorkbookAclRole parseRole(String value) {
        try {
            WorkbookAclRole role = WorkbookAclRole.valueOf(value.trim().toUpperCase());
            if (role == WorkbookAclRole.OWNER) throw ServiceException.validation("Guest share role cannot be OWNER");
            return role;
        } catch (IllegalArgumentException error) {
            throw ServiceException.validation("Guest share role must be viewer, commenter or editor");
        }
    }

    private void requireOwner(String unitId, String actor) {
        if (authorization.role(unitId, actor).orElse(null) != WorkbookAclRole.OWNER) {
            throw ServiceException.forbidden("Workbook role OWNER is required");
        }
    }

    private String randomSecret() {
        byte[] bytes = new byte[32];
        random.nextBytes(bytes);
        return Base64.getUrlEncoder().withoutPadding().encodeToString(bytes);
    }

    private String hash(String value) {
        return HexFormatHolder.format(hashBytes(value));
    }

    private byte[] hashBytes(String value) {
        try {
            return MessageDigest.getInstance("SHA-256").digest(value.getBytes(StandardCharsets.UTF_8));
        } catch (NoSuchAlgorithmException error) {
            throw new IllegalStateException("SHA-256 is unavailable", error);
        }
    }

    private static final class HexFormatHolder {
        private static String format(byte[] value) {
            return java.util.HexFormat.of().formatHex(value);
        }
    }

    public record GuestIdentity(String subject, UUID shareId, String unitId, WorkbookAclRole role, Instant expiresAt) {
    }
}
