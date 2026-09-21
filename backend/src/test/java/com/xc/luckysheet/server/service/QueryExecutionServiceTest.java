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
                    Map.of("local", new QuerySource("sqlite", "jdbc:sqlite:" + file, null, null, null, Map.of()))
            );
            QueryExecutionService service = new QueryExecutionService(properties, access, lifecycle, store, audit, new ObjectMapper());
            var response = service.execute("unit-1", new QueryExecutionRequest(
                    "query-1", "Items", "sqlite", "local", "SELECT name, amount FROM items WHERE amount > ?",
                    null, null, List.of(new com.fasterxml.jackson.databind.node.IntNode(1)), List.of()
            ), "editor");
            assertEquals(List.of("name", "amount"), response.columns());
            assertEquals(1, response.rowCount());
            assertEquals("a", response.rows().get(0).get(0).asText());
            assertEquals(4, response.sourceRevision());
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
                    Map.of("local", new QuerySource("sqlite", "jdbc:sqlite:" + file, null, null, null, Map.of()))
            );
            QueryExecutionService service = new QueryExecutionService(properties, access, lifecycle, store, audit, new ObjectMapper());
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
                    Map.of("local", new QuerySource("sqlite", "jdbc:sqlite:" + file, null, null, null, Map.of()))
            );
            QueryExecutionService service = new QueryExecutionService(properties, access, lifecycle, store, audit, new ObjectMapper());
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
