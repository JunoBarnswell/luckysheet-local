package com.xc.luckysheet.server.migration;

import org.flywaydb.core.Flyway;
import org.h2.tools.RunScript;
import org.junit.jupiter.api.Test;

import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.HexFormat;
import java.util.UUID;

import static org.junit.jupiter.api.Assertions.*;

/** SQL cutover gate tests; these do not substitute for the offline native importer proof tests. */
class NativeKernelCutoverMigrationTest {
    private static final String SNAPSHOT = "{\"version\":10,\"name\":\"存量工作簿\"}";
    private static final String MANIFEST = "{\"version\":11,\"unitId\":\"book\",\"revision\":1,\"pages\":[]}";
    private static final String ENVELOPE = "{\"mutations\":[],\"revision\":1}";

    @Test
    void emptyDatabasePassesAllMigrationsWithoutLegacyRuntimeColumns() throws Exception {
        String url = databaseUrl();
        flyway(url).migrate();
        try (Connection connection = DriverManager.getConnection(url, "sa", "")) {
            assertFalse(columnExists(connection, "WORKBOOKS", "SNAPSHOT_JSON"));
            assertFalse(tableExists(connection, "SNAPSHOT_CHECKPOINT"));
            assertTrue(tableExists(connection, "WORKBOOK_V10_MIGRATION_ARCHIVE"));
            assertFalse(columnExists(connection, "WORKBOOKS", "STORAGE_LOCATION"));
            assertFalse(columnExists(connection, "WORKBOOK_USER_STATE", "DEFAULT_CREATE_LOCATION"));
            assertFalse(columnExists(connection, "WORKBOOK_USER_STATE", "OFFLINE_CACHE"));
            assertFalse(columnExists(connection, "USER_PREFERENCE", "OFFLINE_CACHE"));
            assertTrue(columnExists(connection, "WORKBOOK_USER_STATE", "AUTO_SAVE"));
            assertTrue(columnExists(connection, "WORKBOOK_USER_STATE", "AUTO_SYNC"));
            assertTrue(columnExists(connection, "USER_PREFERENCE", "AUTO_SAVE"));
            assertTrue(columnExists(connection, "USER_PREFERENCE", "AUTO_SYNC"));
        }
    }

    @Test
    void unverifiedWorkbookStopsBeforeAnyDropAndRetainsExactUtf8Sources() throws Exception {
        try (Connection connection = archivedDatabase()) {
            var error = assertThrows(SQLException.class, () -> run(connection, "V12__require_verified_kernel_cutover.sql"));
            assertTrue(error.getMessage().contains("WORKBOOK_V11_MIGRATION_REQUIRED_CK"));
            assertTrue(columnExists(connection, "WORKBOOKS", "SNAPSHOT_JSON"));
            assertTrue(tableExists(connection, "SNAPSHOT_CHECKPOINT"));
            assertEquals(SNAPSHOT, scalar(connection, "select snapshot_json from workbook_v10_migration_archive where source_kind = 'WORKBOOK'"));
            assertEquals(sha256(SNAPSHOT), scalar(connection, "select source_checksum from workbook_v10_migration_archive where source_kind = 'WORKBOOK'"));
            assertEquals(ENVELOPE, scalar(connection, "select envelope_json from workbook_v10_operation_archive"));
        }
    }

    @Test
    void matchingOfflineAttestationAllowsCutoverAndKeepsBothArchives() throws Exception {
        try (Connection connection = archivedDatabase()) {
            insertAttestation(connection);
            run(connection, "V12__require_verified_kernel_cutover.sql");
            assertFalse(columnExists(connection, "WORKBOOKS", "SNAPSHOT_JSON"));
            assertFalse(columnExists(connection, "WORKBOOKS", "SNAPSHOT_REVISION"));
            assertFalse(tableExists(connection, "SNAPSHOT_CHECKPOINT"));
            assertEquals("2", scalar(connection, "select count(*) from workbook_v10_migration_archive"));
            assertEquals("1", scalar(connection, "select count(*) from workbook_v10_operation_archive"));
        }
    }

    @Test
    void changedSourceOrManifestCannotReuseAnEarlierAttestation() throws Exception {
        for (String mutation : new String[] {
                "update workbooks set snapshot_json = '{}'",
                "update workbook_manifests set manifest_json = '{}'",
                "update workbook_v10_migration_archive set snapshot_json = '{}' where source_kind = 'CHECKPOINT'",
                "update operation_log set envelope_json = '{}'",
                "update workbook_v11_migration_proofs set history_verified = false",
                "update workbook_v11_migration_proofs set pages_verified = false"
        }) {
            try (Connection connection = archivedDatabase()) {
                insertAttestation(connection);
                try (var statement = connection.createStatement()) { statement.executeUpdate(mutation); }
                assertThrows(SQLException.class, () -> run(connection, "V12__require_verified_kernel_cutover.sql"), mutation);
                assertTrue(columnExists(connection, "WORKBOOKS", "SNAPSHOT_JSON"), mutation);
                assertTrue(tableExists(connection, "SNAPSHOT_CHECKPOINT"), mutation);
            }
        }
    }

    @Test
    void originalDocumentWithoutRevisionBoundFileReferencePreventsCutover() throws Exception {
        try (Connection connection = archivedDatabase()) {
            insertAttestation(connection);
            byte[] originalBytes = "SQL-gate-binary-source".getBytes(StandardCharsets.UTF_8);
            try (var source = connection.prepareStatement("insert into workbook_native_artifact_migration_archive (unit_id,file_name,mime_type,checksum,byte_length,content,native_metadata_json,created_at,updated_at) values ('book','source.xlsx','application/octet-stream',?,?,?,'{}',current_timestamp,current_timestamp)")) {
                source.setString(1, sha256("SQL-gate-binary-source"));
                source.setInt(2, originalBytes.length); source.setBytes(3, originalBytes); source.executeUpdate();
            }
            assertThrows(SQLException.class, () -> run(connection, "V12__require_verified_kernel_cutover.sql"));
            assertTrue(columnExists(connection, "WORKBOOKS", "SNAPSHOT_JSON"));
            assertTrue(tableExists(connection, "SNAPSHOT_CHECKPOINT"));
        }
    }

    private static Connection archivedDatabase() throws Exception {
        String url = databaseUrl();
        Flyway.configure().dataSource(url, "sa", "").locations("classpath:db/migration/h2").target("9").load().migrate();
        Connection connection = DriverManager.getConnection(url, "sa", "");
        try (var workbook = connection.prepareStatement("insert into workbooks (unit_id,name,snapshot_json,snapshot_revision,revision,created_at,updated_at) values ('book','Existing workbook',?,1,1,current_timestamp,current_timestamp)");
             var checkpoint = connection.prepareStatement("insert into snapshot_checkpoint (unit_id,revision,snapshot_json,checksum,created_at) values ('book',1,?,?,current_timestamp)");
             var operation = connection.prepareStatement("insert into operation_log (operation_id,unit_id,revision,actor_subject,client_sequence,base_revision,envelope_json,committed_at) values ('op-1','book',1,'owner',1,0,?,current_timestamp)")) {
            workbook.setString(1, SNAPSHOT); workbook.executeUpdate();
            checkpoint.setString(1, SNAPSHOT); checkpoint.setString(2, sha256(SNAPSHOT)); checkpoint.executeUpdate();
            operation.setString(1, ENVELOPE); operation.executeUpdate();
        }
        run(connection, "V10__archive_snapshot_migration_boundary.sql");
        return connection;
    }

    /** Supplies gate input only. A real importer must prove native pages and history before writing this row. */
    private static void insertAttestation(Connection connection) throws Exception {
        try (var manifest = connection.prepareStatement("insert into workbook_manifests (manifest_id,unit_id,revision,manifest_version,manifest_json,checksum,created_at) values ('manifest-1','book',1,11,?,?,current_timestamp)");
             var proof = connection.prepareStatement("insert into workbook_v11_migration_proofs (unit_id,proof_version,source_revision,source_snapshot_revision,source_checksum,manifest_checksum,checkpoint_count,operation_count,history_checksum,pages_verified,history_verified,verified_at) values ('book',1,1,1,?,?,1,1,?,true,true,current_timestamp)")) {
            manifest.setString(1, MANIFEST); manifest.setString(2, sha256(MANIFEST)); manifest.executeUpdate();
            proof.setString(1, sha256(SNAPSHOT)); proof.setString(2, sha256(MANIFEST));
            proof.setString(3, sha256("test-only SQL attestation input")); proof.executeUpdate();
        }
    }

    private static void run(Connection connection, String name) throws Exception {
        try (var stream = NativeKernelCutoverMigrationTest.class.getResourceAsStream("/db/migration/h2/" + name)) {
            assertNotNull(stream, name);
            RunScript.execute(connection, new InputStreamReader(stream, StandardCharsets.UTF_8));
        }
    }

    private static boolean tableExists(Connection connection, String table) throws SQLException {
        try (var rows = connection.getMetaData().getTables(null, "PUBLIC", table, null)) { return rows.next(); }
    }

    private static boolean columnExists(Connection connection, String table, String column) throws SQLException {
        try (var rows = connection.getMetaData().getColumns(null, "PUBLIC", table, column)) { return rows.next(); }
    }

    private static String scalar(Connection connection, String sql) throws SQLException {
        try (var statement = connection.createStatement(); var rows = statement.executeQuery(sql)) {
            assertTrue(rows.next()); return rows.getString(1);
        }
    }

    private static String sha256(String value) throws Exception {
        return HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(value.getBytes(StandardCharsets.UTF_8)));
    }

    private static String databaseUrl() { return "jdbc:h2:mem:kernel_cutover_" + UUID.randomUUID() + ";DB_CLOSE_DELAY=-1"; }
    private static Flyway flyway(String url) { return Flyway.configure().dataSource(url, "sa", "").locations("classpath:db/migration/h2").load(); }
}
