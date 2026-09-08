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
import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.sql.Connection;
import java.sql.DriverManager;
import java.time.Duration;
import java.time.Instant;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.atomic.AtomicReference;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
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

    @Test
    void activeQueriesWithColonInDifferentUnitAndQueryIdsDoNotCollide() throws Exception {
        MockRestTransport transport = new MockRestTransport(2);
        ExecutorService callers = Executors.newFixedThreadPool(2);
        QueryExecutionService service = restService(transport.client());
        try {
            Future<?> first = callers.submit(() -> service.execute("tenant:one", restRequest("query"), "editor"));
            Future<?> second = callers.submit(() -> service.execute("tenant", restRequest("one:query"), "editor"));

            assertTrue(transport.requests().await(5, TimeUnit.SECONDS));
            transport.release();
            assertEquals(1, ((com.xc.luckysheet.server.contract.QueryExecutionResponse) first.get(5, TimeUnit.SECONDS)).rowCount());
            assertEquals(1, ((com.xc.luckysheet.server.contract.QueryExecutionResponse) second.get(5, TimeUnit.SECONDS)).rowCount());
        } finally {
            transport.release();
            service.close();
            callers.shutdownNow();
        }
    }

    @Test
    void cancellationCannotCrossTenantThroughColonKeyCollision() throws Exception {
        MockRestTransport transport = new MockRestTransport(1);
        ExecutorService caller = Executors.newSingleThreadExecutor();
        QueryExecutionService service = restService(transport.client());
        try {
            Future<?> running = caller.submit(() -> service.execute("tenant:one", restRequest("query"), "editor"));
            assertTrue(transport.requests().await(5, TimeUnit.SECONDS));

            service.cancel("tenant", "one:query", "editor");
            transport.release();

            assertEquals(1, ((com.xc.luckysheet.server.contract.QueryExecutionResponse) running.get(5, TimeUnit.SECONDS)).rowCount());
        } finally {
            transport.release();
            service.close();
            caller.shutdownNow();
        }
    }

    private QueryExecutionService restService(HttpClient client) throws Exception {
        QueryExecutionProofService proofs = mock(QueryExecutionProofService.class);
        when(proofs.begin(any(String.class), any(String.class), any(String.class), any(Duration.class)))
                .thenAnswer(invocation -> new QueryExecutionProofService.StartedExecution(
                        "execution-" + invocation.getArgument(1), 4));
        when(proofs.publish(any(), any(), any(), any(), any())).thenReturn("sealed-result-hash");
        AccessControlService access = mock(AccessControlService.class);
        when(access.require(any(String.class), any(String.class), eq(WorkbookAclRole.EDITOR)))
                .thenReturn(WorkbookAclRole.EDITOR);
        WorkbookLifecycleService lifecycle = mock(WorkbookLifecycleService.class);
        AuditRecorder audit = mock(AuditRecorder.class);
        QueryProperties properties = new QueryProperties(
                true, 100, 20, 1_000_000, Duration.ofSeconds(5), 2,
                Map.of("rest", new QuerySource("rest", null, null, null,
                        "http://127.0.0.1/", Map.of()))
        );
        QueryExecutionService service = new QueryExecutionService(properties, access, lifecycle, proofs, audit, new ObjectMapper());
        Field http = QueryExecutionService.class.getDeclaredField("http");
        http.setAccessible(true);
        @SuppressWarnings("unchecked")
        AtomicReference<HttpClient> reference = (AtomicReference<HttpClient>) http.get(service);
        reference.set(client);
        return service;
    }

    private QueryExecutionRequest restRequest(String queryId) {
        return new QueryExecutionRequest(queryId, "REST query", "rest", "rest", "data",
                "GET", null, List.of(), List.of());
    }

    private static final class MockRestTransport {
        private final CountDownLatch requests;
        private final List<CompletableFuture<HttpResponse<InputStream>>> responses = new CopyOnWriteArrayList<>();
        private final HttpClient client = mock(HttpClient.class);

        private MockRestTransport(int expectedRequests) {
            requests = new CountDownLatch(expectedRequests);
            when(client.sendAsync(any(HttpRequest.class), any(HttpResponse.BodyHandler.class))).thenAnswer(invocation -> {
                CompletableFuture<HttpResponse<InputStream>> response = new CompletableFuture<>();
                responses.add(response);
                requests.countDown();
                return response;
            });
        }

        private HttpClient client() { return client; }
        private CountDownLatch requests() { return requests; }

        private void release() {
            responses.forEach(pending -> {
                HttpResponse<InputStream> response = mock(HttpResponse.class);
                when(response.statusCode()).thenReturn(200);
                when(response.body()).thenReturn(new ByteArrayInputStream(
                        "[{\"value\":1}]".getBytes(java.nio.charset.StandardCharsets.UTF_8)));
                pending.complete(response);
            });
        }
    }
}
