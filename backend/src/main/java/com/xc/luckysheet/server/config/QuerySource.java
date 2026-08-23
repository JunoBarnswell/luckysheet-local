package com.xc.luckysheet.server.config;

import java.util.Map;

/** Server-side source configuration. Never serialized into a workbook contract. */
public record QuerySource(
        String kind,
        String url,
        String username,
        String password,
        String baseUrl,
        Map<String, String> headers
) {
    public QuerySource {
        if (kind == null || kind.isBlank()) throw new IllegalStateException("Query source kind is required");
        headers = headers == null ? Map.of() : Map.copyOf(headers);
    }
}
