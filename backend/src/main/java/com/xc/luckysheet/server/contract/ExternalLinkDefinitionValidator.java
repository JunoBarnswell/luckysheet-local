package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.service.ServiceException;
import java.util.HashSet;
import java.util.Set;

public final class ExternalLinkDefinitionValidator {
    private ExternalLinkDefinitionValidator() {}
    public static void validateCollection(JsonNode links) {
        if (links == null || links.isNull()) return;
        if (!links.isArray() || links.size() > 100) throw ServiceException.validation("External link collection is invalid");
        Set<String> ids = new HashSet<>(), tokens = new HashSet<>();
        for (JsonNode link : links) {
            validate(link);
            if (!ids.add(link.path("id").asText()) || !tokens.add(link.path("token").asText().toUpperCase(java.util.Locale.ROOT))) throw ServiceException.validation("External link identities must be unique");
        }
    }
    public static void validate(JsonNode link) {
        if (link == null || !link.isObject()) throw ServiceException.validation("External link definition is required");
        link.fieldNames().forEachRemaining(key -> { if (!Set.of("id", "token", "sourceUnitId", "sheets").contains(key)) throw ServiceException.validation("External link definition has unsupported fields"); });
        for (String field : Set.of("id", "token", "sourceUnitId")) if (!link.path(field).isTextual() || !link.path(field).asText().matches("[A-Za-z0-9._:-]{1,200}")) throw ServiceException.validation("External link " + field + " is invalid");
        if (!link.path("sheets").isArray() || link.path("sheets").isEmpty() || link.path("sheets").size() > 1000) throw ServiceException.validation("External sheet bindings are invalid");
        Set<String> aliases = new HashSet<>(), identities = new HashSet<>();
        for (JsonNode sheet : link.path("sheets")) {
            if (!sheet.isObject() || sheet.size() != 2 || !sheet.path("token").isTextual() || sheet.path("token").asText().isBlank() || sheet.path("token").asText().length() > 255
                || !sheet.path("sheetId").isTextual() || !sheet.path("sheetId").asText().matches("[A-Za-z0-9._:-]{1,200}")
                || !aliases.add(sheet.path("token").asText().toUpperCase(java.util.Locale.ROOT)) || !identities.add(sheet.path("sheetId").asText())) throw ServiceException.validation("External sheet binding identity is invalid");
        }
    }
}
