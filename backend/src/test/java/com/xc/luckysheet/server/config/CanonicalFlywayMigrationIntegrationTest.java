package com.xc.luckysheet.server.config;

import org.flywaydb.core.Flyway;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.DriverManager;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.*;

class CanonicalFlywayMigrationIntegrationTest {
    @TempDir Path foreignMigrations;

    private String database() {
        return "jdbc:h2:mem:migration-boundary-" + UUID.randomUUID() + ";DB_CLOSE_DELAY=-1";
    }

    private Flyway canonical(String url) {
        return Flyway.configure().dataSource(url, "sa", "")
                .locations("classpath:db/migration/h2").load();
    }

    @Test
    void migratesAndReopensCanonicalDatabaseWithoutChangingHistory() throws Exception {
        String url = database();
        Flyway flyway = canonical(url);
        CanonicalFlywayMigrationConfiguration.migrate(flyway);
        assertEquals("10", flyway.info().current().getVersion().getVersion());
        int historyCount = flyway.info().applied().length;
        CanonicalFlywayMigrationConfiguration.migrate(canonical(url));
        assertEquals(historyCount, flyway.info().applied().length);
        try (var connection = DriverManager.getConnection(url, "sa", "");
             var statement = connection.createStatement();
             var rows = statement.executeQuery("select source_revision from workbook_source_artifact")) {
            assertFalse(rows.next());
        }
    }

    @Test
    void rejectsFutureDatabaseBeforeAnyRepairAndPreservesItsData() throws Exception {
        String url = database();
        CanonicalFlywayMigrationConfiguration.migrate(canonical(url));
        Files.writeString(foreignMigrations.resolve("V999__different_runtime.sql"),
                "create table future_owned_data(id integer primary key, content varchar(100));"
                        + "insert into future_owned_data values(1, 'must remain intact');");
        Flyway.configure().dataSource(url, "sa", "")
                .locations("classpath:db/migration/h2", "filesystem:" + foreignMigrations).load().migrate();
        Flyway current = canonical(url);
        int historyCount = current.info().applied().length;

        IllegalStateException error = assertThrows(IllegalStateException.class,
                () -> CanonicalFlywayMigrationConfiguration.migrate(current));
        assertTrue(error.getMessage().contains("DATABASE_MIGRATION_INCOMPATIBLE"));
        assertTrue(error.getMessage().contains("999"));
        assertEquals(historyCount, current.info().applied().length);
        try (var connection = DriverManager.getConnection(url, "sa", "");
             var statement = connection.createStatement();
             var rows = statement.executeQuery("select content from future_owned_data where id=1")) {
            assertTrue(rows.next());
            assertEquals("must remain intact", rows.getString(1));
        }
    }

    @Test
    void rejectsChangedMigrationChecksumWithoutRepairingIt() throws Exception {
        String url = database();
        CanonicalFlywayMigrationConfiguration.migrate(canonical(url));
        try (var connection = DriverManager.getConnection(url, "sa", "");
             var statement = connection.createStatement()) {
            assertEquals(1, statement.executeUpdate("update \"flyway_schema_history\" set \"checksum\"=123 where \"version\"='10'"));
        }
        IllegalStateException error = assertThrows(IllegalStateException.class,
                () -> CanonicalFlywayMigrationConfiguration.migrate(canonical(url)));
        assertTrue(error.getMessage().contains("DATABASE_MIGRATION_INCOMPATIBLE"));
        try (var connection = DriverManager.getConnection(url, "sa", "");
             var statement = connection.createStatement();
             var rows = statement.executeQuery("select \"checksum\" from \"flyway_schema_history\" where \"version\"='10'")) {
            assertTrue(rows.next());
            assertEquals(123, rows.getInt(1));
        }
    }
}
