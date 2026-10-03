package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.config.QueryProperties;
import com.xc.luckysheet.server.config.QuerySource;
import com.xc.luckysheet.server.contract.QueryExecutionRequest;
import com.xc.luckysheet.server.contract.QueryStep;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.contract.WorkbookLifecycle;
import com.xc.luckysheet.server.store.WorkbookRow;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.DriverManager;
import java.time.Duration;
import java.time.Instant;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class QueryExecutionServiceTest {
    @Test void restSourceRejectsPathEncodingsAndOversizedBodiesBeforeMaterialization() throws Exception {
        var server = com.sun.net.httpserver.HttpServer.create(new java.net.InetSocketAddress("127.0.0.1", 0), 0);
        var calls = new java.util.concurrent.atomic.AtomicInteger();
        server.createContext("/trusted/api/", exchange -> {
            calls.incrementAndGet();
            byte[] bytes = (exchange.getRequestURI().getPath().endsWith("large") ? "[{\"v\":\"" + "x".repeat(2048) + "\"}]" : "[{\"v\":1}]").getBytes(java.nio.charset.StandardCharsets.UTF_8);
            exchange.sendResponseHeaders(200, bytes.length);
            try (var output = exchange.getResponseBody()) { output.write(bytes); }
        });
        server.start();
        var store = mock(WorkbookStore.class);
        when(store.find("unit-1")).thenReturn(Optional.of(new WorkbookRow("unit-1", "test", "{}", 0, 0, WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now())));
        var properties = new QueryProperties(true, 100, 20, 1024, 2048, Duration.ofSeconds(5), 1,
                Map.of("api", new QuerySource("rest", null, null, null, "http://127.0.0.1:" + server.getAddress().getPort() + "/trusted/api", Map.of("Authorization", "ephemeral-test-only"), java.util.Set.of("unit-1"), java.util.Set.of("editor"))));
        var service = new QueryExecutionService(properties, mock(AccessControlService.class), mock(WorkbookLifecycleService.class), store, mock(WorkbookDataBlockService.class), mock(AuditRecorder.class), new ObjectMapper());
        try {
            var response = service.execute("unit-1", new QueryExecutionRequest("q", "Items", "rest", "api", "items", "GET", null, List.of(), List.of()), "editor");
            assertEquals(1, response.rows().getFirst().getFirst().asInt());
            for (String path : List.of("../escape", "/escape", "%2e%2e/escape", "%252e%252e/escape", "//example.invalid/escape")) {
                assertThrows(ServiceException.class, () -> service.execute("unit-1", new QueryExecutionRequest("denied", "Denied", "rest", "api", path, "GET", null, List.of(), List.of()), "editor"));
            }
            assertEquals(1, calls.get());
            assertThrows(ServiceException.class, () -> service.execute("unit-1", new QueryExecutionRequest("large", "Large", "rest", "api", "large", "GET", null, List.of(), List.of()), "editor"));
        } finally { service.close(); server.stop(0); }
    }
    @Test
    void sqliteQueryRunsOnServerWithConfiguredSourceAndNoClientCredentials() throws Exception {
        Path file = Files.createTempFile("luckysheet-query-", ".db");
        try (Connection connection = DriverManager.getConnection("jdbc:sqlite:" + file)) {
            connection.createStatement().execute("CREATE TABLE items(name TEXT, amount INTEGER)");
            connection.createStatement().execute("INSERT INTO items VALUES ('a', 3), ('b', 1)");
        }
        try {
            WorkbookStore store = mock(WorkbookStore.class);
            when(store.find("unit-1")).thenReturn(Optional.of(new WorkbookRow("unit-1", "test", "{}", 0, 4,
                    WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now())));
            AccessControlService access = mock(AccessControlService.class);
            when(access.require("unit-1", "editor", WorkbookAclRole.EDITOR)).thenReturn(WorkbookAclRole.EDITOR);
            WorkbookLifecycleService lifecycle = mock(WorkbookLifecycleService.class);
            AuditRecorder audit = mock(AuditRecorder.class);
            QueryProperties properties = new QueryProperties(
                    true, 100, 20, 1_000_000, 2_048, Duration.ofSeconds(5), 2,
                    Map.of("local", new QuerySource("sqlite", "jdbc:sqlite:" + file, null, null, null, Map.of(), java.util.Set.of(), java.util.Set.of("editor", "owner")))
            );
            QueryExecutionService service = new QueryExecutionService(properties, access, lifecycle, store, mock(WorkbookDataBlockService.class), audit, new ObjectMapper());
            var response = service.execute("unit-1", new QueryExecutionRequest(
                    "query-1", "Items", "sqlite", "local", "SELECT name, amount FROM items WHERE amount > ?",
                    null, null, List.of(new com.fasterxml.jackson.databind.node.IntNode(1)), List.of()
            ), "editor");
            assertEquals(List.of("name", "amount"), response.columns());
            assertEquals(1, response.rowCount());
            assertEquals("a", response.rows().get(0).get(0).asText());
            assertEquals(4, response.sourceRevision());
            for (String statement : List.of("SELECT pg_terminate_backend(42)", "SELECT \"pg_terminate_backend\"(42)", "SELECT evil.lower('x')", "SELECT [randomblob](1000000000)", "SELECT `randomblob`(1000000000)", "SELECT \"as\"()", "SELECT sum$unknown()", "SELECT evil#comment\n()", "SELECT 1; DELETE FROM items", "WITH x AS (SELECT 1) DELETE FROM items")) {
                assertThrows(ServiceException.class, () -> service.execute("unit-1", new QueryExecutionRequest("denied", "Denied", "sqlite", "local", statement, null, null, List.of(), List.of()), "editor"));
            }
            var count = service.execute("unit-1", new QueryExecutionRequest("count", "Count", "sqlite", "local", "WITH x(a) AS (SELECT amount FROM items) SELECT COUNT(a) FROM x", null, null, List.of(), List.of()), "editor");
            assertEquals(2, count.rows().getFirst().getFirst().asInt());
            var shortBudget = new QueryProperties(true, 100, 20, 1000000, 2048, Duration.ofMillis(100), 1, properties.sources());
            var timeoutService = new QueryExecutionService(shortBudget, access, lifecycle, store, mock(WorkbookDataBlockService.class), audit, new ObjectMapper());
            try {
                org.junit.jupiter.api.Assertions.assertTimeout(Duration.ofSeconds(2), () -> assertThrows(ServiceException.class, () -> timeoutService.execute("unit-1", new QueryExecutionRequest("slow", "Slow", "sqlite", "local", "WITH RECURSIVE x(a) AS (VALUES(1) UNION ALL SELECT a+1 FROM x WHERE a<100000000) SELECT SUM(a) FROM x", null, null, List.of(), List.of()), "editor")));
                assertEquals(2, timeoutService.execute("unit-1", new QueryExecutionRequest("control", "Control", "sqlite", "local", "SELECT COUNT(*) FROM items", null, null, List.of(), List.of()), "editor").rows().getFirst().getFirst().asInt());
            } finally { timeoutService.close(); }

            assertThrows(ServiceException.class, () -> service.execute("unit-1", new QueryExecutionRequest("unauthorized", "Denied", "sqlite", "local", "SELECT name FROM items", null, null, List.of(), List.of()), "unbound-subject"));

            service.close();
        } finally {
            Files.deleteIfExists(file);
        }
    }

    @Test
    void serverQueryReplaysPersistedCleaningRecipeSteps() throws Exception {
        Path file = Files.createTempFile("luckysheet-query-recipe-", ".db");
        try (Connection connection = DriverManager.getConnection("jdbc:sqlite:" + file)) {
            connection.createStatement().execute("CREATE TABLE items(name TEXT, amount INTEGER)");
            connection.createStatement().execute("INSERT INTO items VALUES ('  a,one  ', 3), ('  a,one  ', 8), ('b,two', 1)");
        }
        try {
            WorkbookStore store = mock(WorkbookStore.class);
            when(store.find("unit-recipe")).thenReturn(Optional.of(new WorkbookRow("unit-recipe", "test", "{}", 0, 1,
                    WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now())));
            AccessControlService access = mock(AccessControlService.class);
            when(access.require("unit-recipe", "editor", WorkbookAclRole.EDITOR)).thenReturn(WorkbookAclRole.EDITOR);
            WorkbookLifecycleService lifecycle = mock(WorkbookLifecycleService.class);
            AuditRecorder audit = mock(AuditRecorder.class);
            QueryProperties properties = new QueryProperties(
                    true, 100, 20, 1_000_000, 2_048, Duration.ofSeconds(5), 2,
                    Map.of("local", new QuerySource("sqlite", "jdbc:sqlite:" + file, null, null, null, Map.of(), java.util.Set.of(), java.util.Set.of("editor", "owner")))
            );
            QueryExecutionService service = new QueryExecutionService(properties, access, lifecycle, store, mock(WorkbookDataBlockService.class), audit, new ObjectMapper());
            ObjectMapper mapper = new ObjectMapper();
            var response = service.execute("unit-recipe", new QueryExecutionRequest(
                    "query-recipe", "Recipe", "sqlite", "local", "SELECT name, amount FROM items",
                    null, null, List.of(), List.of(
                            new QueryStep("trim", "trim-text", "Trim name", mapper.readTree("{\"columns\":[\"name\"]}"), true),
                            new QueryStep("dedupe", "remove-duplicates", "Dedupe name", mapper.readTree("{\"columns\":[\"name\"]}"), true),
                            new QueryStep("split", "split-column", "Split name", mapper.readTree("{\"column\":\"name\",\"delimiter\":\",\",\"outputColumns\":[\"first\",\"second\"]}"), true)
                    )
            ), "editor");
            assertEquals(List.of("first", "second", "amount"), response.columns());
            assertEquals(2, response.rowCount());
            assertEquals("a", response.rows().get(0).get(0).asText());
            assertEquals("one", response.rows().get(0).get(1).asText());
            service.close();
        } finally {
            Files.deleteIfExists(file);
        }
    }

    @Test
    void blockQuerySessionReturnsBoundedPagesAndRejectsGaps() throws Exception {
        Path file = Files.createTempFile("luckysheet-query-blocks-", ".db");
        try (Connection connection = DriverManager.getConnection("jdbc:sqlite:" + file)) {
            connection.createStatement().execute("CREATE TABLE items(name TEXT, amount INTEGER)");
            for (int index = 0; index < 5; index += 1) {
                connection.createStatement().execute("INSERT INTO items VALUES ('item-" + index + "', " + index + ")");
            }
        }
        try {
            WorkbookStore store = mock(WorkbookStore.class);
            when(store.find("unit-blocks")).thenReturn(Optional.of(new WorkbookRow("unit-blocks", "test", "{}", 0, 2,
                    WorkbookLifecycle.ACTIVE, Instant.now(), Instant.now())));
            AccessControlService access = mock(AccessControlService.class);
            when(access.require("unit-blocks", "editor", WorkbookAclRole.EDITOR)).thenReturn(WorkbookAclRole.EDITOR);
            when(access.require("unit-blocks", "editor", WorkbookAclRole.VIEWER)).thenReturn(WorkbookAclRole.EDITOR);
            WorkbookLifecycleService lifecycle = mock(WorkbookLifecycleService.class);
            AuditRecorder audit = mock(AuditRecorder.class);
            QueryProperties properties = new QueryProperties(
                    true, 100, 20, 1_000_000, 2_048, Duration.ofSeconds(5), 2,
                    Map.of("local", new QuerySource("sqlite", "jdbc:sqlite:" + file, null, null, null, Map.of(), java.util.Set.of(), java.util.Set.of("editor", "owner")))
            );
            QueryExecutionService service = new QueryExecutionService(properties, access, lifecycle, store, mock(WorkbookDataBlockService.class), audit, new ObjectMapper());
            var request = new QueryExecutionRequest(
                    "query-blocks", "Items", "sqlite", "local", "SELECT name, amount FROM items ORDER BY amount",
                    null, null, List.of(), List.of()
            );
            var session = service.executeBlocks("unit-blocks", request, "editor");
            assertEquals(5, session.rowCount());
            assertEquals(2_048, session.blockRowCount());
            assertEquals(List.of("text", "number"), session.columnTypes());
            assertEquals(5, service.readBlock("unit-blocks", request.queryId(), session.executionId(), 0, "editor").rows().size());
            assertThrows(ServiceException.class, () -> service.readBlock("unit-blocks", request.queryId(), session.executionId(), 1, "editor"));
            service.finishBlocks("unit-blocks", request.queryId(), session.executionId(), "editor");
            assertThrows(ServiceException.class, () -> service.readBlock("unit-blocks", request.queryId(), session.executionId(), 0, "editor"));
            service.close();
        } finally {
            Files.deleteIfExists(file);
        }
    }
}
