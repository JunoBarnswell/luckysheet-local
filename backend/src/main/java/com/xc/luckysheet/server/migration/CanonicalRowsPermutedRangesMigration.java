package com.xc.luckysheet.server.migration;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.flywaydb.core.api.migration.BaseJavaMigration;
import org.flywaydb.core.api.migration.Context;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/** One-way clean-break migration for the rows.permuted committed-operation contract. */
public abstract class CanonicalRowsPermutedRangesMigration extends BaseJavaMigration {
    private static final int MAX_ROW = 1_048_575;
    private static final int MAX_COLUMN = 16_383;
    private final ObjectMapper mapper = new ObjectMapper();

    @Override
    public void migrate(Context context) throws Exception {
        rewrite(context.getConnection(),
                "select operation_id, envelope_json from operation_log",
                "update operation_log set envelope_json = ? where operation_id = ?",
                "operation_log");
        rewrite(context.getConnection(),
                "select event_id, payload_json from coordination_outbox where published_at is null",
                "update coordination_outbox set payload_json = ? where event_id = ?",
                "coordination_outbox");
    }

    private void rewrite(Connection connection, String query, String update, String source) throws Exception {
        try (PreparedStatement select = connection.prepareStatement(query);
             ResultSet rows = select.executeQuery();
             PreparedStatement write = connection.prepareStatement(update)) {
            while (rows.next()) {
                String id = rows.getString(1);
                String payload = rows.getString(2);
                ObjectNode envelope = parse(source, id, payload);
                if (!canonicalize(envelope, source, id)) continue;
                write.setString(1, mapper.writeValueAsString(envelope));
                write.setString(2, id);
                if (write.executeUpdate() != 1) throw new IllegalStateException(source + " migration lost record " + id);
            }
        }
    }

    private ObjectNode parse(String source, String id, String payload) throws Exception {
        JsonNode parsed = mapper.readTree(payload);
        if (!(parsed instanceof ObjectNode envelope)) {
            throw new IllegalStateException(source + " record " + id + " is not an operation envelope");
        }
        return envelope;
    }

    private boolean canonicalize(ObjectNode envelope, String source, String id) {
        JsonNode mutationsNode = envelope.get("mutations");
        if (!(mutationsNode instanceof ArrayNode mutations)) {
            throw new IllegalStateException(source + " record " + id + " has no mutations array");
        }
        boolean changed = false;
        for (JsonNode entry : mutations) {
            if (!(entry instanceof ObjectNode mutation)) throw new IllegalStateException(source + " record " + id + " has malformed mutation");
            if (!"rows.permuted".equals(mutation.path("id").asText())) continue;
            String sheetId = requiredText(mutation, "sheetId", source, id);
            JsonNode paramsNode = mutation.get("params");
            if (!(paramsNode instanceof ObjectNode params)) throw new IllegalStateException(source + " record " + id + " has rows.permuted without params");
            JsonNode rangeNode = params.get("range");
            if (!(rangeNode instanceof ObjectNode range)) throw new IllegalStateException(source + " record " + id + " has rows.permuted without range");
            if (!sheetId.equals(requiredText(range, "sheetId", source, id))) {
                throw new IllegalStateException(source + " record " + id + " has rows.permuted range on another sheet");
            }
            int startRow = coordinate(range, "startRow", source, id);
            int endRow = coordinate(range, "endRow", source, id);
            if (endRow < startRow) throw new IllegalStateException(source + " record " + id + " has inverted rows.permuted range");
            ArrayNode ranges = mutation.putArray("affectedRanges");
            ObjectNode canonical = ranges.addObject();
            canonical.put("sheetId", sheetId);
            canonical.put("startRow", startRow);
            canonical.put("endRow", endRow);
            canonical.put("startColumn", 0);
            canonical.put("endColumn", MAX_COLUMN);
            changed = true;
        }
        return changed;
    }

    private static String requiredText(ObjectNode object, String field, String source, String id) {
        JsonNode value = object.get(field);
        if (value == null || !value.isTextual() || value.asText().isBlank()) {
            throw new IllegalStateException(source + " record " + id + " has invalid " + field);
        }
        return value.asText();
    }

    private static int coordinate(ObjectNode range, String field, String source, String id) {
        JsonNode value = range.get(field);
        if (value == null || !value.isIntegralNumber() || !value.canConvertToInt() || value.intValue() < 0 || value.intValue() > MAX_ROW) {
            throw new IllegalStateException(source + " record " + id + " has invalid " + field);
        }
        return value.intValue();
    }
}
