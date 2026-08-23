package com.xc.luckysheet.server.service;

import java.nio.charset.StandardCharsets;
import java.util.Base64;

/** Shared bounded page cursor codec for application query services and HTTP. */
public final class CursorPageRequest {
    private static final int DEFAULT_LIMIT = 20;
    private static final int MAX_LIMIT = 50;

    private CursorPageRequest() {}

    public static int limit(Integer requested) {
        if (requested == null) return DEFAULT_LIMIT;
        if (requested < 1 || requested > MAX_LIMIT) throw ServiceException.validation("limit must be between 1 and " + MAX_LIMIT);
        return requested;
    }

    public static int page(String cursor) {
        if (cursor == null || cursor.isBlank()) return 0;
        try {
            String decoded = new String(Base64.getUrlDecoder().decode(cursor), StandardCharsets.UTF_8);
            int value = Integer.parseInt(decoded);
            if (value < 0) throw new NumberFormatException();
            return value;
        } catch (IllegalArgumentException error) {
            throw ServiceException.validation("cursor is invalid");
        }
    }

    public static String next(int page) {
        return Base64.getUrlEncoder().withoutPadding().encodeToString(Integer.toString(page).getBytes(StandardCharsets.UTF_8));
    }
}
