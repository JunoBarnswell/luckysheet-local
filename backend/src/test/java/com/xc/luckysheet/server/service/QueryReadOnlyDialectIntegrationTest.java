package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.config.QueryProperties;
import com.xc.luckysheet.server.config.QuerySource;
import com.xc.luckysheet.server.contract.QueryExecutionRequest;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.store.WorkbookRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.condition.EnabledIfEnvironmentVariable;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.*;
import java.time.Duration;
import java.time.Instant;
import java.util.*;
import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

/** Optional real-database acceptance; credentials are supplied outside the repository. */
@EnabledIfEnvironmentVariable(named = "REMEDIATION_JDBC_CONFIG", matches = ".+")
class QueryReadOnlyDialectIntegrationTest {
    @Test void postgresEnforcesReadOnlyAndRejectsSessionControlCalls() throws Exception { verifyDialect("postgresql"); }
    @Test void mysqlEnforcesReadOnlyAndRejectsSessionControlCalls() throws Exception { verifyDialect("mysql"); }
    private void verifyDialect(String dialect) throws Exception {
        var config = new Properties();
        try (var input = Files.newInputStream(Path.of(System.getenv("REMEDIATION_JDBC_CONFIG")))) { config.load(input); }
        String url = config.getProperty(dialect + ".url"), username = config.getProperty(dialect + ".username"), password = config.getProperty(dialect + ".password");
        try (var connection = DriverManager.getConnection(url, username, password); var statement = connection.createStatement()) {
            statement.execute("DROP TABLE IF EXISTS items"); statement.execute("CREATE TABLE items(amount INTEGER)"); statement.execute("INSERT INTO items VALUES (7)");
            connection.setReadOnly(true); connection.setAutoCommit(false);
            statement.execute(dialect.equals("postgresql") ? "SET TRANSACTION READ ONLY" : "START TRANSACTION READ ONLY");
            assertThrows(SQLException.class, () -> statement.execute("INSERT INTO items VALUES (99)"));
            connection.rollback();
        }
        var store = mock(WorkbookStore.class);
        when(store.find("book")).thenReturn(Optional.of(new WorkbookRow("book", "Book", "{}", 0, 0, WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now())));
        var source = new QuerySource("jdbc", url, username, password, null, Map.of(), Set.of("book"), Set.of("editor"));
        var properties = new QueryProperties(true, 100, 20, 100000, 2048, Duration.ofSeconds(5), 1, Map.of("db", source));
        var service = new QueryExecutionService(properties, mock(AccessControlService.class), mock(WorkbookLifecycleService.class), store, mock(WorkbookDataBlockService.class), mock(AuditRecorder.class), new ObjectMapper());
        try {
            var response = service.execute("book", request("SELECT amount FROM items"), "editor");
            assertEquals(7, response.rows().getFirst().getFirst().asInt());
            for (String sql : List.of("SELECT pg_terminate_backend(42)", "SELECT pg_catalog.pg_try_advisory_lock(42)", "SELECT GET_LOCK('test',0)", "SELECT SLEEP(1)", "SELECT untrusted_function()", "WITH x AS (SELECT 1) DELETE FROM items")) assertThrows(ServiceException.class, () -> service.execute("book", request(sql), "editor"));
            assertEquals(1, service.execute("book", request("SELECT COUNT(*) FROM items"), "editor").rows().getFirst().getFirst().asInt());
        } finally { service.close(); }
    }
    private QueryExecutionRequest request(String sql) { return new QueryExecutionRequest("q", "Query", "jdbc", "db", sql, null, null, List.of(), List.of()); }
}
