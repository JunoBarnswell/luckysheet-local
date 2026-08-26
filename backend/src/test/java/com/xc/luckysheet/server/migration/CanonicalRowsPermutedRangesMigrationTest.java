package com.xc.luckysheet.server.migration;

import org.flywaydb.core.api.migration.Context;
import org.junit.jupiter.api.Test;

import java.sql.DriverManager;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class CanonicalRowsPermutedRangesMigrationTest {
    @Test
    void rewritesOperationLogAndPendingOutboxToTheCanonicalFullColumnRange() throws Exception {
        try (var connection = DriverManager.getConnection("jdbc:h2:mem:rows_permuted_migration;DB_CLOSE_DELAY=-1")) {
            try (var statement = connection.createStatement()) {
                statement.execute("create table operation_log (operation_id varchar(200) primary key, envelope_json text not null)");
                statement.execute("create table coordination_outbox (event_id varchar(36) primary key, payload_json text not null, published_at timestamp with time zone)");
                String payload = envelope("sheet-1", 4, 9);
                statement.executeUpdate("insert into operation_log values ('op-1', '" + payload.replace("'", "''") + "')");
                statement.executeUpdate("insert into coordination_outbox values ('event-1', '" + payload.replace("'", "''") + "', null)");
            }
            Context context = mock(Context.class);
            when(context.getConnection()).thenReturn(connection);

            new db.migration.h2.V6__canonical_rows_permuted_ranges().migrate(context);

            try (var statement = connection.createStatement(); var rows = statement.executeQuery("select envelope_json from operation_log where operation_id = 'op-1'")) {
                rows.next();
                assertTrue(rows.getString(1).contains("\"startColumn\":0"));
                assertTrue(rows.getString(1).contains("\"endColumn\":16383"));
            }
            try (var statement = connection.createStatement(); var rows = statement.executeQuery("select payload_json from coordination_outbox where event_id = 'event-1'")) {
                rows.next();
                assertTrue(rows.getString(1).contains("\"endColumn\":16383"));
            }
        }
    }

    @Test
    void rejectsMalformedHistoricalRowsInsteadOfRetainingTheLegacyRange() throws Exception {
        try (var connection = DriverManager.getConnection("jdbc:h2:mem:rows_permuted_migration_invalid;DB_CLOSE_DELAY=-1")) {
            try (var statement = connection.createStatement()) {
                statement.execute("create table operation_log (operation_id varchar(200) primary key, envelope_json text not null)");
                statement.execute("create table coordination_outbox (event_id varchar(36) primary key, payload_json text not null, published_at timestamp with time zone)");
                statement.executeUpdate("insert into operation_log values ('op-invalid', '{\"mutations\":[{\"id\":\"rows.permuted\",\"sheetId\":\"sheet-1\",\"params\":{}}]}')");
            }
            Context context = mock(Context.class);
            when(context.getConnection()).thenReturn(connection);

            assertThrows(IllegalStateException.class, () -> new db.migration.h2.V6__canonical_rows_permuted_ranges().migrate(context));
        }
    }

    private static String envelope(String sheetId, int startRow, int endRow) {
        return "{\"schema\":\"OperationEnvelope\",\"mutations\":[{\"id\":\"rows.permuted\",\"sheetId\":\"" + sheetId
                + "\",\"params\":{\"sheetId\":\"" + sheetId + "\",\"range\":{\"sheetId\":\"" + sheetId
                + "\",\"startRow\":" + startRow + ",\"endRow\":" + endRow + ",\"startColumn\":1,\"endColumn\":2}},\"affectedRanges\":[{\"sheetId\":\"" + sheetId
                + "\",\"startRow\":" + startRow + ",\"endRow\":" + endRow + ",\"startColumn\":1,\"endColumn\":2}]}]}";
    }
}
