package com.xc.luckysheet.server.migration;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import org.flywaydb.core.api.migration.BaseJavaMigration;
import org.flywaydb.core.api.migration.Context;
import java.sql.Connection;

/** Explicit migration boundary: historical operations belong to one retired session. */
public abstract class ClientSessionMigration extends BaseJavaMigration {
    @Override public Integer getChecksum() { return 1; }
    @Override public void migrate(Context context) throws Exception {
        Connection connection = context.getConnection();
        validate(connection, "operation_log", "envelope_json");
        validate(connection, "coordination_outbox", "payload_json");
        try (var statement = connection.createStatement()) {
            statement.execute("alter table operation_log add column client_session_id varchar(200) not null default 'migrated-session'");
            boolean mysql = connection.getMetaData().getDatabaseProductName().toLowerCase().contains("mysql");
            statement.execute("alter table operation_log drop " + (mysql ? "index " : "constraint ") + "operation_log_unit_actor_sequence_uk");
            statement.execute("alter table operation_log add constraint operation_log_unit_actor_session_sequence_uk unique(unit_id, actor_subject, client_session_id, client_sequence)");
        }
        rewrite(connection, "operation_log", "operation_id", "envelope_json");
        rewrite(connection, "coordination_outbox", "event_id", "payload_json");
        try (var statement = connection.createStatement()) {
            statement.execute("alter table operation_log alter column client_session_id drop default");
        }
    }
    private void validate(Connection connection, String table, String column) throws Exception {
        var mapper = new ObjectMapper();
        try (var query = connection.prepareStatement("select " + column + " from " + table); var rows = query.executeQuery()) {
            while (rows.next()) {
                var value = mapper.readTree(rows.getString(1));
                if (!(value instanceof ObjectNode object) || !"OperationEnvelope".equals(object.path("schema").asText()) || object.has("clientSessionId")) {
                    throw new IllegalStateException("SESSION_MIGRATION_INVALID_ENVELOPE: " + table);
                }
            }
        }
    }
    private void rewrite(Connection connection, String table, String key, String column) throws Exception {
        var mapper = new ObjectMapper();
        try (var query = connection.prepareStatement("select " + key + ", " + column + " from " + table);
             var rows = query.executeQuery();
             var update = connection.prepareStatement("update " + table + " set " + column + "=? where " + key + "=?")) {
            while (rows.next()) {
                var value = mapper.readTree(rows.getString(2));
                if (!(value instanceof ObjectNode object) || !"OperationEnvelope".equals(object.path("schema").asText())) {
                    throw new IllegalStateException("SESSION_MIGRATION_INVALID_ENVELOPE: " + table + "/" + rows.getString(1));
                }
                if (object.has("clientSessionId")) throw new IllegalStateException("SESSION_MIGRATION_UNEXPECTED_VERSION: " + rows.getString(1));
                object.put("clientSessionId", "migrated-session");
                update.setString(1, mapper.writeValueAsString(object)); update.setString(2, rows.getString(1));
                if (update.executeUpdate() != 1) throw new IllegalStateException("SESSION_MIGRATION_WRITE_FAILED");
            }
        }
    }
}
