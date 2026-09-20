package com.xc.luckysheet.server.migration;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.flywaydb.core.api.migration.Context;
import org.junit.jupiter.api.Test;
import java.sql.DriverManager;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

class ClientSessionMigrationTest {
    @Test void migratesHistoricalEnvelopeAndScopesSequenceToSession() throws Exception {
        try (var connection = DriverManager.getConnection("jdbc:h2:mem:sessions-" + java.util.UUID.randomUUID())) {
            try (var statement = connection.createStatement()) {
                statement.execute("create table operation_log(operation_id varchar primary key, unit_id varchar, actor_subject varchar, client_sequence bigint, envelope_json varchar, constraint operation_log_unit_actor_sequence_uk unique(unit_id,actor_subject,client_sequence))");
                statement.execute("create table coordination_outbox(event_id varchar primary key, payload_json varchar)");
                statement.execute("insert into operation_log values('old','book','user',1,'{\"schema\":\"OperationEnvelope\"}')");
                statement.execute("insert into coordination_outbox values('event','{\"schema\":\"OperationEnvelope\"}')");
            }
            Context context = mock(Context.class); when(context.getConnection()).thenReturn(connection);
            new db.migration.h2.V8__client_session().migrate(context);
            try (var statement = connection.createStatement(); var rows = statement.executeQuery("select client_session_id,envelope_json from operation_log")) {
                assertTrue(rows.next()); assertEquals("migrated-session", rows.getString(1));
                assertEquals("migrated-session", new ObjectMapper().readTree(rows.getString(2)).path("clientSessionId").asText());
            }
            try (var statement = connection.createStatement()) {
                statement.execute("insert into operation_log values('tab-a','book','user',1,'{}','tab-a')");
                statement.execute("insert into operation_log values('tab-b','book','user',1,'{}','tab-b')");
                assertThrows(java.sql.SQLException.class, () -> statement.execute("insert into operation_log values('duplicate','book','user',1,'{}','tab-a')"));
                assertThrows(java.sql.SQLException.class, () -> statement.execute("insert into operation_log(operation_id,unit_id,actor_subject,client_sequence,envelope_json) values('missing','book','user',2,'{}')"));
            }
        }
    }
    @Test void unknownEnvelopeBlocksMigration() throws Exception {
        try (var connection = DriverManager.getConnection("jdbc:h2:mem:invalid-" + java.util.UUID.randomUUID())) {
            try (var statement = connection.createStatement()) {
                statement.execute("create table operation_log(operation_id varchar primary key, unit_id varchar, actor_subject varchar, client_sequence bigint, envelope_json varchar, constraint operation_log_unit_actor_sequence_uk unique(unit_id,actor_subject,client_sequence))");
                statement.execute("create table coordination_outbox(event_id varchar primary key, payload_json varchar)");
                statement.execute("insert into operation_log values('bad','book','user',1,'{\"schema\":\"Unknown\"}')");
            }
            Context context = mock(Context.class); when(context.getConnection()).thenReturn(connection);
            assertThrows(IllegalStateException.class, () -> new db.migration.h2.V8__client_session().migrate(context));
        }
    }
}
