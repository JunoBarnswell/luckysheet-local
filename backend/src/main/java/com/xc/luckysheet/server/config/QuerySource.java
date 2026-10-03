package com.xc.luckysheet.server.config;

import java.util.Map;

/** Server-side source configuration. Never serialized into a workbook contract. */
public record QuerySource(
        String kind,
        String url,
        String username,
        String password,
        String baseUrl,
        Map<String, String> headers,
        java.util.Set<String> allowedWorkbooks,
        java.util.Set<String> allowedSubjects
) {
    public QuerySource {
        if (kind == null || kind.isBlank()) throw new IllegalStateException("Query source kind is required");
        headers = headers == null ? Map.of() : Map.copyOf(headers);
        allowedWorkbooks = allowedWorkbooks == null ? java.util.Set.of() : java.util.Set.copyOf(allowedWorkbooks);
        allowedSubjects = allowedSubjects == null ? java.util.Set.of() : java.util.Set.copyOf(allowedSubjects);
    }
    public void requireAccess(String unitId, String subject) {
        if ((allowedWorkbooks.isEmpty() && allowedSubjects.isEmpty())
                || (!allowedWorkbooks.isEmpty() && !allowedWorkbooks.contains(unitId))
                || (!allowedSubjects.isEmpty() && !allowedSubjects.contains(subject))) {
            throw com.xc.luckysheet.server.service.ServiceException.forbidden("The configured query source is not granted to this workbook and subject");
        }
    }
}
