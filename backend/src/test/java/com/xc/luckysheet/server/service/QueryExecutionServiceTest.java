package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.config.QueryProperties;
import com.xc.luckysheet.server.config.QuerySource;
import com.xc.luckysheet.server.contract.QueryExecutionRequest;
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
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.eq;

class QueryExecutionServiceTest {
    @Test
    void sqliteQueryRunsOnServerWithConfiguredSourceAndNoClientCredentials() throws Exception {
        Path file = Files.createTempFile("luckysheet-query-", ".db");
        try (Connection connection = DriverManager.getConnection("jdbc:sqlite:" + file)) {
            connection.createStatement().execute("CREATE TABLE items(name TEXT, amount INTEGER)");
            connection.createStatement().execute("INSERT INTO items VALUES ('a', 3), ('b', 1)");
        }
        try {
            QueryExecutionProofService proofs = mock(QueryExecutionProofService.class);
            when(proofs.begin(eq("unit-1"), eq("query-1"), eq("editor"), any()))
                    .thenReturn(new QueryExecutionProofService.StartedExecution("execution-1", 4));
            when(proofs.publish(eq("unit-1"), eq("query-1"), eq("execution-1"), eq("editor"), any()))
                    .thenReturn("sealed-result-hash");
            AccessControlService access = mock(AccessControlService.class);
            when(access.require("unit-1", "editor", WorkbookAclRole.EDITOR)).thenReturn(WorkbookAclRole.EDITOR);
            WorkbookLifecycleService lifecycle = mock(WorkbookLifecycleService.class);
            AuditRecorder audit = mock(AuditRecorder.class);
            QueryProperties properties = new QueryProperties(
                    true, 100, 20, 1_000_000, Duration.ofSeconds(5), 2,
                    Map.of("local", new QuerySource("sqlite", "jdbc:sqlite:" + file, null, null, null, Map.of()))
            );
            QueryExecutionService service = new QueryExecutionService(properties, access, lifecycle, proofs, audit, new ObjectMapper());
            var response = service.execute("unit-1", new QueryExecutionRequest(
                    "query-1", "Items", "sqlite", "local", "SELECT name, amount FROM items WHERE amount > ?",
                    null, null, List.of(new com.fasterxml.jackson.databind.node.IntNode(1)), List.of()
            ), "editor");
            assertEquals(List.of("name", "amount"), response.columns());
            assertEquals(1, response.rowCount());
            assertEquals("a", response.rows().get(0).get(0).asText());
            assertEquals(4, response.sourceRevision());
            assertEquals("execution-1", response.executionToken());
            assertEquals("sealed-result-hash", response.resultHash());
            service.close();
        } finally {
            Files.deleteIfExists(file);
        }
    }
}
