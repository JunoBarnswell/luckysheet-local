package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.JsonNodeFactory;
import com.fasterxml.jackson.databind.node.ArrayNode;
import com.xc.luckysheet.server.config.QueryProperties;
import com.xc.luckysheet.server.config.QuerySource;
import com.xc.luckysheet.server.contract.QueryExecutionRequest;
import com.xc.luckysheet.server.contract.QueryExecutionResponse;
import com.xc.luckysheet.server.contract.QueryStep;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;

import jakarta.annotation.PreDestroy;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Future;
import java.util.concurrent.FutureTask;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.RejectedExecutionException;
import java.util.concurrent.CancellationException;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.concurrent.atomic.AtomicReference;

@Service
public class QueryExecutionService {
    private static final Set<String> STEP_KINDS = Set.of("source", "filter", "select-columns", "rename-column", "sort", "group-by", "join", "pivot");
    private static final Set<String> FILTER_OPERATORS = Set.of("eq", "neq", "contains", "startsWith", "endsWith", "gt", "gte", "lt", "lte", "isNull", "notNull");
    private static final Set<String> AGGREGATIONS = Set.of("sum", "count", "average", "min", "max");
    private static final Set<String> WRITE_SQL_KEYWORDS = Set.of(
            "insert", "update", "delete", "drop", "alter", "create", "truncate", "grant", "revoke", "copy", "call", "merge", "replace",
            "into", "outfile", "dumpfile", "load", "set", "reset", "use", "lock", "unlock", "attach", "detach", "vacuum", "analyze", "pragma",
            "reindex", "refresh", "comment", "execute", "prepare", "deallocate"
    );

    private final QueryProperties properties;
    private final AccessControlService access;
    private final WorkbookLifecycleService lifecycle;
    private final WorkbookStore store;
    private final AuditRecorder audit;
    private final ObjectMapper mapper;
    private final ExecutorService workers;
    private final HttpClient http;
    private final Map<String, ActiveQuery> active = new ConcurrentHashMap<>();

    public QueryExecutionService(
            QueryProperties properties,
            AccessControlService access,
            WorkbookLifecycleService lifecycle,
            WorkbookStore store,
            AuditRecorder audit,
            ObjectMapper mapper
    ) {
        this.properties = properties;
        this.access = access;
        this.lifecycle = lifecycle;
        this.store = store;
        this.audit = audit;
        this.mapper = mapper;
        this.workers = new ThreadPoolExecutor(properties.workerThreads(), properties.workerThreads(), 0, TimeUnit.MILLISECONDS,
                new ArrayBlockingQueue<>(Math.max(8, properties.workerThreads() * 8)), runnable -> {
            Thread thread = new Thread(runnable, "server-query-worker");
            thread.setDaemon(true);
            return thread;
        }, new ThreadPoolExecutor.AbortPolicy());
        this.http = HttpClient.newBuilder().connectTimeout(properties.timeout()).build();
    }

    public QueryExecutionResponse execute(String unitId, QueryExecutionRequest request, String actor) {
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        lifecycle.requireActive(unitId);
        if (!properties.enabled()) throw ServiceException.unavailable("Server query execution is disabled");
        QuerySource source;
        try {
            source = properties.requireSource(request.sourceRef(), request.connectorId());
            validateRequest(request, source);
        } catch (IllegalArgumentException error) {
            audit.rejected(request.queryId(), unitId, actor, "QUERY_EXECUTION", error.getMessage());
            throw ServiceException.validation(error.getMessage());
        }

        Instant started = Instant.now();
        ExecutionControl control = new ExecutionControl();
        FutureTask<QueryTable> future = new FutureTask<>(() -> executeInternal(request, source, control));
        String executionKey = unitId + ":" + request.queryId();
        ActiveQuery running = new ActiveQuery(actor, control, future);
        ActiveQuery previous = active.putIfAbsent(executionKey, running);
        if (previous != null) {
            control.cancel();
            future.cancel(true);
            throw ServiceException.conflict("A query with this id is already running");
        }
        try {
            workers.execute(future);
        } catch (RejectedExecutionException error) {
            active.remove(executionKey, running);
            control.cancel();
            future.cancel(true);
            throw ServiceException.unavailable("Query execution queue is full");
        }
        try {
            QueryTable table = future.get(properties.timeout().toMillis(), TimeUnit.MILLISECONDS);
            checkFinalSize(table);
            long duration = Duration.between(started, Instant.now()).toMillis();
            audit.accepted(request.queryId(), unitId, actor, "QUERY_EXECUTION", null, mapper.createObjectNode()
                    .put("connectorId", request.connectorId())
                    .put("sourceRef", request.sourceRef())
                    .put("rowCount", table.rows.size())
                    .put("durationMs", duration));
            long sourceRevision = store.find(unitId).map(row -> row.revision()).orElse(0L);
            return new QueryExecutionResponse(request.queryId(), request.connectorId(), request.sourceRef(), sourceRevision,
                    table.columns, table.rows, table.rows.size(), Instant.now(), duration);
        } catch (TimeoutException error) {
            control.cancel();
            future.cancel(true);
            audit.rejected(request.queryId(), unitId, actor, "QUERY_EXECUTION", "Query timed out");
            throw ServiceException.timeout("Query timed out");
        } catch (InterruptedException error) {
            control.cancel();
            future.cancel(true);
            Thread.currentThread().interrupt();
            audit.rejected(request.queryId(), unitId, actor, "QUERY_EXECUTION", "Query was cancelled");
            throw ServiceException.timeout("Query was cancelled");
        } catch (CancellationException error) {
            audit.rejected(request.queryId(), unitId, actor, "QUERY_EXECUTION", "Query was cancelled");
            throw ServiceException.timeout("Query was cancelled");
        } catch (ExecutionException error) {
            Throwable cause = error.getCause() == null ? error : error.getCause();
            if (control.cancelled() || cause instanceof CancellationException) {
                audit.rejected(request.queryId(), unitId, actor, "QUERY_EXECUTION", "Query was cancelled");
                throw ServiceException.timeout("Query was cancelled");
            }
            String reason = cause instanceof QueryFailure failure ? failure.safeMessage() : "Query execution failed";
            audit.rejected(request.queryId(), unitId, actor, "QUERY_EXECUTION", reason);
            if (cause instanceof QueryFailure failure) throw failure.exception();
            throw ServiceException.validation(reason);
        } finally {
            active.remove(executionKey, running);
        }
    }

    public void cancel(String unitId, String queryId, String actor) {
        access.require(unitId, actor, WorkbookAclRole.EDITOR);
        lifecycle.requireActive(unitId);
        ActiveQuery query = active.get(unitId + ":" + queryId);
        if (query == null) throw ServiceException.notFound("Running query not found");
        if (!query.actor().equals(actor) && !access.currentRole(unitId, actor).includes(WorkbookAclRole.OWNER)) {
            throw ServiceException.forbidden("Only the query owner or workbook owner may cancel a query");
        }
        query.control().cancel();
        query.future().cancel(true);
        audit.accepted(queryId, unitId, actor, "QUERY_CANCEL", null, mapper.createObjectNode());
    }

    @PreDestroy
    public void close() {
        workers.shutdownNow();
    }

    private QueryTable executeInternal(QueryExecutionRequest request, QuerySource source, ExecutionControl control) {
        control.ensureActive();
        QueryTable sourceTable = switch (request.connectorId().toLowerCase(Locale.ROOT)) {
            case "jdbc", "sqlite" -> executeJdbc(request, source, control);
            case "rest" -> executeRest(request, source, control);
            default -> throw QueryFailure.validation("Only server JDBC, SQLite and REST connectors are executable");
        };
        checkSize(sourceTable);
        QueryTable current = sourceTable;
        for (QueryStep step : request.steps()) {
            control.ensureActive();
            if (!step.enabled() || step.kind().equals("source")) continue;
            if (!STEP_KINDS.contains(step.kind())) throw QueryFailure.validation("Unsupported query step kind: " + step.kind());
            current = applyStep(current, step, control);
            checkSize(current);
        }
        return current;
    }

    private QueryTable executeJdbc(QueryExecutionRequest request, QuerySource source, ExecutionControl control) {
        if (source.url() == null || source.url().isBlank()) throw QueryFailure.validation("Configured JDBC source has no URL");
        if (!source.url().toLowerCase(Locale.ROOT).startsWith("jdbc:")) throw QueryFailure.validation("Configured source URL must be a JDBC URL");
        if (request.connectorId().equalsIgnoreCase("sqlite") && !source.url().toLowerCase(Locale.ROOT).startsWith("jdbc:sqlite:")) {
            throw QueryFailure.validation("SQLite source must use a jdbc:sqlite URL");
        }
        String sql = readOnlySql(request.statement());
        try (Connection connection = DriverManager.getConnection(jdbcReadOnlyUrl(request, source), nullToEmpty(source.username()), nullToEmpty(source.password()))) {
            control.ensureActive();
            try {
                boolean sqliteUriReadOnly = request.connectorId().equalsIgnoreCase("sqlite");
                if (!sqliteUriReadOnly) connection.setReadOnly(true);
                connection.setAutoCommit(false);
                if (!sqliteUriReadOnly && !connection.isReadOnly()) throw QueryFailure.validation("Read-only database session was not accepted by the driver");
            } catch (SQLException error) {
                throw QueryFailure.validation("Read-only database session is unavailable");
            }
            try (PreparedStatement statement = connection.prepareStatement(sql, ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY)) {
                control.bind(statement);
                statement.setQueryTimeout(timeoutSeconds());
                statement.setFetchSize(QueryTable.CHUNK_ROWS);
                for (int index = 0; index < request.parameters().size(); index++) bind(statement, index + 1, request.parameters().get(index));
                control.ensureActive();
                try (ResultSet result = statement.executeQuery()) {
                    return readResult(result, control);
                } finally {
                    control.unbind(statement);
                }
            } finally {
                try { connection.rollback(); } catch (SQLException ignored) { }
            }
        } catch (SQLException error) {
            if (control.cancelled() || Thread.currentThread().isInterrupted()) throw QueryFailure.cancelled();
            throw QueryFailure.validation("Configured database query failed");
        }
    }

    private QueryTable executeRest(QueryExecutionRequest request, QuerySource source, ExecutionControl control) {
        if (source.baseUrl() == null || source.baseUrl().isBlank()) throw QueryFailure.validation("Configured REST source has no base URL");
        URI target;
        try {
            URI base = URI.create(source.baseUrl());
            if (request.statement().startsWith("http://") || request.statement().startsWith("https://")) {
                throw QueryFailure.validation("REST statement must be a path relative to sourceRef");
            }
            target = base.resolve(request.statement());
            if (!sameOrigin(base, target)) throw QueryFailure.validation("REST path leaves the configured source origin");
        } catch (Exception error) {
            if (error instanceof QueryFailure failure) throw failure;
            throw QueryFailure.validation("REST source URL is invalid");
        }
        try {
            HttpRequest.Builder builder = HttpRequest.newBuilder(target).timeout(properties.timeout());
            source.headers().forEach(builder::header);
            String method = request.method() == null || request.method().isBlank() ? "GET" : request.method().toUpperCase(Locale.ROOT);
            if (method.equals("POST")) {
                String body = request.body() == null ? "{}" : mapper.writeValueAsString(request.body());
                builder.header("content-type", "application/json").POST(HttpRequest.BodyPublishers.ofString(body, StandardCharsets.UTF_8));
            } else {
                builder.GET();
            }
            CompletableFuture<HttpResponse<InputStream>> pending = http.sendAsync(builder.build(), HttpResponse.BodyHandlers.ofInputStream());
            control.bind(pending);
            HttpResponse<InputStream> response;
            try {
                response = pending.get();
            } catch (InterruptedException error) {
                Thread.currentThread().interrupt();
                throw QueryFailure.cancelled();
            } catch (ExecutionException error) {
                if (control.cancelled()) throw QueryFailure.cancelled();
                throw error;
            } finally {
                control.unbind(pending);
            }
            if (response.statusCode() < 200 || response.statusCode() >= 300) {
                try (InputStream ignored = response.body()) { }
                throw QueryFailure.validation("REST source returned an unsuccessful status");
            }
            try (InputStream body = response.body()) {
                control.bind(body);
                byte[] bytes;
                try {
                    bytes = readBounded(body, properties.maxResponseBytes(), control);
                } finally {
                    control.unbind(body);
                }
                if (bytes.length > properties.maxResponseBytes()) throw QueryFailure.validation("REST response is too large");
                control.ensureActive();
                JsonNode parsed = mapper.readTree(new String(bytes, StandardCharsets.UTF_8));
                if (parsed == null) throw QueryFailure.validation("REST response is empty");
                return parseRestResponse(parsed);
            }
        } catch (QueryFailure error) {
            throw error;
        } catch (IOException error) {
            if (control.cancelled()) throw QueryFailure.cancelled();
            throw QueryFailure.validation("REST query failed");
        } catch (Exception error) {
            if (control.cancelled()) throw QueryFailure.cancelled();
            throw QueryFailure.validation("REST query failed");
        }
    }

    private byte[] readBounded(InputStream input, int maximumBytes, ExecutionControl control) throws IOException {
        java.io.ByteArrayOutputStream output = new java.io.ByteArrayOutputStream(Math.min(maximumBytes, 64 * 1024));
        byte[] buffer = new byte[Math.min(64 * 1024, maximumBytes + 1)];
        int total = 0;
        while (total <= maximumBytes) {
            control.ensureActive();
            int read = input.read(buffer, 0, Math.min(buffer.length, maximumBytes + 1 - total));
            if (read < 0) break;
            output.write(buffer, 0, read);
            total += read;
            if (total > maximumBytes) break;
        }
        return output.toByteArray();
    }

    private QueryTable parseRestResponse(JsonNode root) {
        JsonNode data = root.isArray() ? root : root.path("data");
        if (data.isArray()) return recordsToTable(data);
        if (root.isObject()) {
            ArrayNode records = mapper.createArrayNode();
            records.add(root);
            return recordsToTable(records);
        }
        throw QueryFailure.validation("REST response must be an object or array of objects");
    }

    /** Opens SQLite through its URI read-only mode instead of trusting SQL text alone. */
    private String jdbcReadOnlyUrl(QueryExecutionRequest request, QuerySource source) {
        if (!request.connectorId().equalsIgnoreCase("sqlite")) return source.url();
        String prefix = "jdbc:sqlite:";
        String value = source.url().substring(prefix.length());
        if (value.equals(":memory:") || value.startsWith("file::memory:")) return source.url();
        if (!value.startsWith("file:")) value = "file:" + value.replace('\\', '/');
        if (value.matches(".*[?&]mode=(?!ro(?:&|$))[^&]+.*")) {
            throw QueryFailure.validation("SQLite source must be opened in read-only mode");
        }
        if (value.matches(".*[?&]mode=ro(?:&|$).*") || value.endsWith("?mode=ro")) return prefix + value;
        return prefix + value + (value.contains("?") ? "&mode=ro" : "?mode=ro");
    }

    private QueryTable readResult(ResultSet result, ExecutionControl control) throws SQLException {
        ResultSetMetaData metadata = result.getMetaData();
        if (metadata.getColumnCount() > properties.maxColumns()) throw QueryFailure.validation("Query returned too many columns");
        List<String> columns = new ArrayList<>();
        for (int index = 1; index <= metadata.getColumnCount(); index++) columns.add(metadata.getColumnLabel(index));
        List<QueryChunk> chunks = new ArrayList<>();
        List<List<JsonNode>> rows = new ArrayList<>(QueryTable.CHUNK_ROWS);
        int totalRows = 0;
        while (result.next()) {
            control.ensureActive();
            if (totalRows >= properties.maxRows()) throw QueryFailure.validation("Query returned too many rows");
            List<JsonNode> row = new ArrayList<>();
            for (int index = 1; index <= metadata.getColumnCount(); index++) row.add(toNode(result.getObject(index)));
            rows.add(row);
            totalRows++;
            if (rows.size() == QueryTable.CHUNK_ROWS) {
                chunks.add(QueryChunk.fromRows(rows, columns.size()));
                rows = new ArrayList<>(QueryTable.CHUNK_ROWS);
            }
        }
        if (!rows.isEmpty()) chunks.add(QueryChunk.fromRows(rows, columns.size()));
        return QueryTable.fromChunks(columns, chunks);
    }

    private QueryTable applyStep(QueryTable input, QueryStep step, ExecutionControl control) {
        return switch (step.kind()) {
            case "filter" -> filter(input, step, control);
            case "select-columns" -> select(input, step, control);
            case "rename-column" -> rename(input, step, control);
            case "sort" -> sort(input, step, control);
            case "group-by" -> group(input, step, control);
            case "join" -> join(input, step, control);
            case "pivot" -> pivot(input, step, control);
            default -> throw QueryFailure.validation("Unsupported query step: " + step.kind());
        };
    }

    private QueryTable filter(QueryTable input, QueryStep step, ExecutionControl control) {
        String column = required(step.config(), "column", step.id());
        int index = input.columnIndex(column, step.id());
        String operator = step.config().path("operator").asText("eq");
        if (!FILTER_OPERATORS.contains(operator)) throw QueryFailure.validation("Unsupported filter operator");
        JsonNode expected = scalarOrNull(step.config().get("value"));
        boolean caseSensitive = !step.config().has("caseSensitive") || step.config().path("caseSensitive").asBoolean();
        List<QueryChunk> chunks = new ArrayList<>();
        for (QueryChunk chunk : input.chunks) {
            List<List<JsonNode>> selected = new ArrayList<>();
            for (int rowIndex = 0; rowIndex < chunk.rowCount(); rowIndex++) {
                control.ensureActive();
                List<JsonNode> row = chunk.row(rowIndex);
                if (matches(row.get(index), expected, operator, caseSensitive)) selected.add(row);
            }
            if (!selected.isEmpty()) chunks.add(QueryChunk.fromRows(selected, input.columns.size()));
        }
        return QueryTable.fromChunks(input.columns, chunks);
    }

    private QueryTable select(QueryTable input, QueryStep step, ExecutionControl control) {
        List<String> columns = stringList(step.config().get("columns"), step.id());
        int[] indexes = columns.stream().mapToInt(name -> input.columnIndex(name, step.id())).toArray();
        List<QueryChunk> chunks = new ArrayList<>();
        for (QueryChunk chunk : input.chunks) {
            control.ensureActive();
            chunks.add(chunk.select(indexes));
        }
        return QueryTable.fromChunks(columns, chunks);
    }

    private QueryTable rename(QueryTable input, QueryStep step, ExecutionControl control) {
        String from = required(step.config(), "from", step.id());
        String to = required(step.config(), "to", step.id());
        int index = input.columnIndex(from, step.id());
        List<String> columns = new ArrayList<>(input.columns);
        if (columns.contains(to) && !from.equals(to)) throw QueryFailure.validation("Renamed column already exists");
        columns.set(index, to);
        control.ensureActive();
        return input.withColumns(columns);
    }

    private QueryTable sort(QueryTable input, QueryStep step, ExecutionControl control) {
        List<SortKey> keys = new ArrayList<>();
        JsonNode by = step.config().get("by");
        if (by != null && by.isArray()) {
            for (JsonNode item : by) keys.add(new SortKey(required(item, "column", step.id()), !item.has("ascending") || item.path("ascending").asBoolean()));
        } else {
            keys.add(new SortKey(required(step.config(), "column", step.id()), !step.config().has("ascending") || step.config().path("ascending").asBoolean()));
        }
        control.ensureActive();
        List<List<JsonNode>> rows = new ArrayList<>(input.rows);
        rows.sort((left, right) -> {
            control.ensureActive();
            for (SortKey key : keys) {
                int comparison = compare(left.get(input.columnIndex(key.column(), step.id())), right.get(input.columnIndex(key.column(), step.id())));
                if (comparison != 0) return key.ascending() ? comparison : -comparison;
            }
            return 0;
        });
        return new QueryTable(input.columns, rows);
    }

    private QueryTable group(QueryTable input, QueryStep step, ExecutionControl control) {
        List<String> groups = stringList(first(step.config(), "by", "columns", "groupBy"), step.id());
        List<Aggregate> aggregates = aggregates(step.config().get("aggregations"), input, step.id());
        int[] indexes = groups.stream().mapToInt(name -> input.columnIndex(name, step.id())).toArray();
        Map<String, List<List<JsonNode>>> buckets = new LinkedHashMap<>();
        for (List<JsonNode> row : input.rows) {
            control.ensureActive();
            String key = indexesAsJson(row, indexes);
            buckets.computeIfAbsent(key, ignored -> new ArrayList<>()).add(row);
        }
        List<List<JsonNode>> rows = new ArrayList<>();
        for (List<List<JsonNode>> bucket : buckets.values()) {
            control.ensureActive();
            List<JsonNode> row = new ArrayList<>();
            List<JsonNode> key = bucket.get(0) == null ? List.of() : indexesAsValues(bucket.get(0), indexes);
            row.addAll(key);
            for (Aggregate aggregate : aggregates) row.add(aggregate(bucket, aggregate, input, step.id()));
            rows.add(row);
        }
        return new QueryTable(concat(groups, aggregates.stream().map(Aggregate::as).toList()), rows);
    }

    private QueryTable join(QueryTable input, QueryStep step, ExecutionControl control) {
        JsonNode rightConfig = step.config().has("right") ? step.config().get("right") : step.config().get("rightTable");
        control.ensureActive();
        QueryTable right = tableFromJson(rightConfig, step.id());
        checkSize(right);
        List<String> leftOn = stringList(first(step.config(), "leftOn", "on"), step.id());
        List<String> rightOn = step.config().has("rightOn") ? stringList(step.config().get("rightOn"), step.id()) : leftOn;
        if (leftOn.size() != rightOn.size()) throw QueryFailure.validation("Join keys must have equal lengths");
        int[] leftIndexes = leftOn.stream().mapToInt(name -> input.columnIndex(name, step.id())).toArray();
        int[] rightIndexes = rightOn.stream().mapToInt(name -> right.columnIndex(name, step.id())).toArray();
        Map<String, List<List<JsonNode>>> matches = new HashMap<>();
        for (List<JsonNode> row : right.rows) {
            control.ensureActive();
            matches.computeIfAbsent(indexesAsJson(row, rightIndexes), ignored -> new ArrayList<>()).add(row);
        }
        String type = step.config().path("type").asText("inner");
        if (!Set.of("inner", "left", "full").contains(type)) throw QueryFailure.validation("Join type is invalid");
        List<List<JsonNode>> rows = new ArrayList<>();
        Set<List<JsonNode>> matched = new HashSet<>();
        for (List<JsonNode> left : input.rows) {
            control.ensureActive();
            List<List<JsonNode>> candidates = matches.getOrDefault(indexesAsJson(left, leftIndexes), List.of());
            if (candidates.isEmpty()) {
                if (type.equals("left") || type.equals("full")) rows.add(concatValues(left, nulls(right.columns.size())));
            } else for (List<JsonNode> rightRow : candidates) { matched.add(rightRow); rows.add(concatValues(left, rightRow)); }
        }
        if (type.equals("full")) for (List<JsonNode> row : right.rows) {
            control.ensureActive();
            if (!matched.contains(row)) rows.add(concatValues(nulls(input.columns.size()), row));
        }
        List<String> rightColumns = right.columns.stream().map(column -> input.columns.contains(column) ? column + "_right" : column).toList();
        return new QueryTable(concat(input.columns, rightColumns), rows);
    }

    private QueryTable pivot(QueryTable input, QueryStep step, ExecutionControl control) {
        List<String> rowFields = stringList(first(step.config(), "rows", "rowFields"), step.id());
        List<String> columnFields = stringList(first(step.config(), "columns", "columnFields"), step.id());
        List<String> values = stringList(first(step.config(), "values", "valueFields"), step.id());
        String function = step.config().path("aggregation").asText("sum");
        if (!AGGREGATIONS.contains(function)) throw QueryFailure.validation("Unsupported pivot aggregation");
        int[] rowIndexes = rowFields.stream().mapToInt(name -> input.columnIndex(name, step.id())).toArray();
        int[] columnIndexes = columnFields.stream().mapToInt(name -> input.columnIndex(name, step.id())).toArray();
        int[] valueIndexes = values.stream().mapToInt(name -> input.columnIndex(name, step.id())).toArray();
        List<String> columnKeys = new ArrayList<>();
        Map<String, String> columnKeyLabels = new LinkedHashMap<>();
        Set<String> seenColumnKeys = new HashSet<>();
        for (List<JsonNode> row : input.rows) {
            control.ensureActive();
            String key = indexesAsJson(row, columnIndexes);
            if (seenColumnKeys.add(key)) {
                columnKeys.add(key);
                columnKeyLabels.put(key, indexesAsValues(row, columnIndexes).stream()
                        .map(this::queryScalarLabel)
                        .collect(java.util.stream.Collectors.joining(" / ")));
            }
        }
        Map<String, List<List<JsonNode>>> groups = new LinkedHashMap<>();
        for (List<JsonNode> row : input.rows) {
            control.ensureActive();
            groups.computeIfAbsent(indexesAsJson(row, rowIndexes), ignored -> new ArrayList<>()).add(row);
        }
        List<String> columns = new ArrayList<>(rowFields);
        for (String key : columnKeys) for (String value : values) {
            String label = columnKeyLabels.get(key) + " · " + value;
            if (columns.contains(label)) throw QueryFailure.validation("Pivot produces duplicate column " + label);
            columns.add(label);
        }
        List<List<JsonNode>> rows = new ArrayList<>();
        for (List<List<JsonNode>> group : groups.values()) {
            control.ensureActive();
            List<JsonNode> row = new ArrayList<>(indexesAsValues(group.get(0), rowIndexes));
            for (String key : columnKeys) {
                List<List<JsonNode>> matching = group.stream().filter(value -> indexesAsJson(value, columnIndexes).equals(key)).toList();
                for (int index : valueIndexes) row.add(aggregateValues(matching.stream().map(value -> value.get(index)).toList(), function));
            }
            rows.add(row);
        }
        return new QueryTable(columns, rows);
    }

    private void validateRequest(QueryExecutionRequest request, QuerySource source) {
        if (!Set.of("jdbc", "sqlite", "rest").contains(request.connectorId().toLowerCase(Locale.ROOT))) {
            throw new IllegalArgumentException("Only jdbc, sqlite and rest connectors are available on the server");
        }
        if (request.connectorId().equalsIgnoreCase("rest") && (source.baseUrl() == null || source.baseUrl().isBlank())) {
            throw new IllegalArgumentException("REST source baseUrl is not configured");
        }
        for (QueryStep step : request.steps()) {
            if (!STEP_KINDS.contains(step.kind())) throw new IllegalArgumentException("Unsupported query step kind: " + step.kind());
        }
    }

    private QueryTable tableFromJson(JsonNode value, String stepId) {
        if (value == null) throw QueryFailure.validation("Join step " + stepId + " requires right data");
        if (value.isArray()) return recordsToTable(value);
        if (value.isObject() && value.path("data").isArray()) return recordsToTable(value.path("data"));
        if (value.isObject() && value.path("columns").isArray() && value.path("rows").isArray()) {
            List<String> columns = stringList(value.path("columns"), stepId);
            List<List<JsonNode>> rows = new ArrayList<>();
            for (JsonNode row : value.path("rows")) {
                if (!row.isArray() || row.size() != columns.size()) throw QueryFailure.validation("Join row width is invalid");
                rows.add(row.elements().hasNext() ? iterableToList(row) : List.of());
            }
            return new QueryTable(columns, rows);
        }
        throw QueryFailure.validation("Join step right data is invalid");
    }

    private QueryTable recordsToTable(JsonNode records) {
        if (!records.isArray()) throw QueryFailure.validation("Records must be an array");
        if (records.size() > properties.maxRows()) throw QueryFailure.validation("Query returned too many rows");
        LinkedHashMap<String, Boolean> columns = new LinkedHashMap<>();
        for (JsonNode record : records) {
            if (!record.isObject()) throw QueryFailure.validation("Record rows must be objects");
            record.fieldNames().forEachRemaining(name -> columns.put(name, true));
            if (columns.size() > properties.maxColumns()) throw QueryFailure.validation("Query returned too many columns");
        }
        List<String> names = new ArrayList<>(columns.keySet());
        List<List<JsonNode>> rows = new ArrayList<>();
        for (JsonNode record : records) rows.add(names.stream().map(name -> scalarOrNull(record.get(name))).toList());
        QueryTable table = new QueryTable(names, rows);
        checkSize(table);
        return table;
    }

    private List<JsonNode> iterableToList(JsonNode array) {
        List<JsonNode> values = new ArrayList<>();
        array.elements().forEachRemaining(value -> values.add(scalarOrNull(value)));
        return values;
    }

    private List<Aggregate> aggregates(JsonNode raw, QueryTable input, String stepId) {
        if (raw == null || !raw.isArray() || raw.isEmpty()) return List.of(new Aggregate("*", "count", "count"));
        List<Aggregate> result = new ArrayList<>();
        for (JsonNode item : raw) {
            String fn = item.path("function").asText(item.path("fn").asText("sum"));
            if (!AGGREGATIONS.contains(fn)) throw QueryFailure.validation("Unsupported aggregation in " + stepId);
            String column = item.path("column").asText("*");
            if (!column.equals("*")) input.columnIndex(column, stepId);
            String as = item.path("as").asText(fn + "_" + column);
            result.add(new Aggregate(column, fn, as));
        }
        return result;
    }

    private JsonNode aggregate(List<List<JsonNode>> rows, Aggregate aggregate, QueryTable input, String stepId) {
        List<JsonNode> values = aggregate.column().equals("*")
                ? rows.stream().<JsonNode>map(ignored -> JsonNodeFactory.instance.numberNode(1)).toList()
                : rows.stream().map(row -> row.get(input.columnIndex(aggregate.column(), stepId))).toList();
        return aggregateValues(values, aggregate.function());
    }

    private JsonNode aggregateValues(List<JsonNode> values, String function) {
        if (function.equals("count")) return JsonNodeFactory.instance.numberNode(values.stream().filter(value -> value != null && !value.isNull()).count());
        List<Double> numbers = values.stream().filter(value -> value != null && value.isNumber()).map(JsonNode::asDouble).toList();
        if (numbers.isEmpty()) return JsonNodeFactory.instance.nullNode();
        return switch (function) {
            case "sum" -> JsonNodeFactory.instance.numberNode(numbers.stream().mapToDouble(Double::doubleValue).sum());
            case "average" -> JsonNodeFactory.instance.numberNode(numbers.stream().mapToDouble(Double::doubleValue).average().orElse(0));
            case "min" -> JsonNodeFactory.instance.numberNode(numbers.stream().mapToDouble(Double::doubleValue).min().orElse(0));
            case "max" -> JsonNodeFactory.instance.numberNode(numbers.stream().mapToDouble(Double::doubleValue).max().orElse(0));
            default -> throw QueryFailure.validation("Unsupported aggregation");
        };
    }

    private boolean matches(JsonNode actual, JsonNode expected, String operator, boolean caseSensitive) {
        JsonNode left = normalize(actual, caseSensitive);
        JsonNode right = normalize(expected, caseSensitive);
        return switch (operator) {
            case "eq" -> left.equals(right);
            case "neq" -> !left.equals(right);
            case "contains" -> left.isTextual() && right.isTextual() && left.asText().contains(right.asText());
            case "startsWith" -> left.isTextual() && right.isTextual() && left.asText().startsWith(right.asText());
            case "endsWith" -> left.isTextual() && right.isTextual() && left.asText().endsWith(right.asText());
            case "gt" -> compare(left, right) > 0;
            case "gte" -> compare(left, right) >= 0;
            case "lt" -> compare(left, right) < 0;
            case "lte" -> compare(left, right) <= 0;
            case "isNull" -> left == null || left.isNull();
            case "notNull" -> left != null && !left.isNull();
            default -> false;
        };
    }

    private JsonNode normalize(JsonNode value, boolean caseSensitive) {
        JsonNode scalar = scalarOrNull(value);
        if (!caseSensitive && scalar != null && scalar.isTextual()) return JsonNodeFactory.instance.textNode(scalar.asText().toLowerCase(Locale.ROOT));
        return scalar;
    }

    private int compare(JsonNode left, JsonNode right) {
        return WorkbookCollationComparator.compare(left, right);
    }

    private List<JsonNode> nulls(int count) {
        return java.util.Collections.nCopies(count, JsonNodeFactory.instance.nullNode());
    }

    private List<JsonNode> concatValues(List<JsonNode> left, List<JsonNode> right) {
        List<JsonNode> values = new ArrayList<>(left);
        values.addAll(right);
        return values;
    }

    private List<String> concat(List<String> left, List<String> right) {
        List<String> values = new ArrayList<>(left);
        values.addAll(right);
        return values;
    }

    private List<String> stringList(JsonNode node, String stepId) {
        if (node == null || !node.isArray() || node.isEmpty() || !allText(node)) throw QueryFailure.validation("Step " + stepId + " requires a non-empty column list");
        List<String> values = new ArrayList<>();
        node.forEach(value -> values.add(value.asText()));
        return values;
    }

    private JsonNode first(JsonNode config, String... names) {
        for (String name : names) if (config.has(name)) return config.get(name);
        return null;
    }

    private String required(JsonNode node, String key, String stepId) {
        String value = node.path(key).asText("");
        if (value.isBlank()) throw QueryFailure.validation("Step " + stepId + " requires " + key);
        return value;
    }

    private boolean allText(JsonNode node) {
        for (JsonNode value : node) if (!value.isTextual() || value.asText().isBlank()) return false;
        return true;
    }

    private String indexesAsJson(List<JsonNode> row, int[] indexes) {
        List<JsonNode> values = indexesAsValues(row, indexes);
        try { return mapper.writeValueAsString(values); } catch (Exception error) { throw new IllegalStateException(error); }
    }

    private List<JsonNode> indexesAsValues(List<JsonNode> row, int[] indexes) {
        List<JsonNode> values = new ArrayList<>();
        for (int index : indexes) values.add(row.get(index));
        return values;
    }

    private String queryScalarLabel(JsonNode value) {
        if (value == null || value.isNull()) return "";
        if (value.isNumber()) {
            java.math.BigDecimal number = value.decimalValue().stripTrailingZeros();
            return number.signum() == 0 ? "0" : number.toPlainString();
        }
        return value.asText();
    }

    private JsonNode scalarOrNull(JsonNode node) {
        if (node == null || node.isNull()) return JsonNodeFactory.instance.nullNode();
        if (!node.isValueNode()) throw QueryFailure.validation("Query values must be scalar");
        return node.deepCopy();
    }

    private JsonNode toNode(Object value) {
        if (value == null) return JsonNodeFactory.instance.nullNode();
        if (value instanceof Boolean bool) return JsonNodeFactory.instance.booleanNode(bool);
        if (value instanceof Number number) return mapper.valueToTree(number);
        if (value instanceof byte[] bytes) return JsonNodeFactory.instance.textNode(java.util.Base64.getEncoder().encodeToString(bytes));
        return JsonNodeFactory.instance.textNode(String.valueOf(value));
    }

    private void bind(PreparedStatement statement, int index, JsonNode value) throws SQLException {
        JsonNode scalar = scalarOrNull(value);
        if (scalar.isNull()) statement.setObject(index, null);
        else if (scalar.isBoolean()) statement.setBoolean(index, scalar.asBoolean());
        else if (scalar.isIntegralNumber()) statement.setLong(index, scalar.asLong());
        else if (scalar.isFloatingPointNumber()) statement.setDouble(index, scalar.asDouble());
        else statement.setString(index, scalar.asText());
    }

    private String readOnlySql(String sql) {
        String normalized = sql.trim();
        List<String> tokens = sqlTokens(normalized);
        if (tokens.isEmpty()) throw QueryFailure.validation("Only one read-only SQL statement is allowed");
        String first = tokens.getFirst();
        if (!(first.equals("select") || first.equals("with") || first.equals("values") || first.equals("explain"))) {
            throw QueryFailure.validation("Only read-only SQL statements are allowed");
        }
        if (tokens.stream().anyMatch(WRITE_SQL_KEYWORDS::contains)) throw QueryFailure.validation("Read-only SQL cannot contain a write operation");
        return normalized;
    }

    /** Tokenizes SQL outside literals/comments; the JDBC read-only session remains the enforcement authority. */
    private List<String> sqlTokens(String sql) {
        List<String> tokens = new ArrayList<>();
        StringBuilder word = new StringBuilder();
        for (int index = 0; index < sql.length(); index++) {
            char current = sql.charAt(index);
            if (current == '-' && index + 1 < sql.length() && sql.charAt(index + 1) == '-') {
                flushToken(word, tokens);
                index += 2;
                while (index < sql.length() && sql.charAt(index) != '\n') index++;
                continue;
            }
            if (current == '/' && index + 1 < sql.length() && sql.charAt(index + 1) == '*') {
                flushToken(word, tokens);
                index += 2;
                boolean closed = false;
                while (index + 1 < sql.length()) {
                    if (sql.charAt(index) == '*' && sql.charAt(index + 1) == '/') { closed = true; index++; break; }
                    index++;
                }
                if (!closed) throw QueryFailure.validation("SQL comment is not closed");
                continue;
            }
            if (current == '\'' || current == '"' || current == '`') {
                flushToken(word, tokens);
                char quote = current;
                boolean closed = false;
                while (++index < sql.length()) {
                    if (sql.charAt(index) == quote) {
                        if (index + 1 < sql.length() && sql.charAt(index + 1) == quote) { index++; continue; }
                        closed = true;
                        break;
                    }
                }
                if (!closed) throw QueryFailure.validation("SQL literal is not closed");
                continue;
            }
            if (current == ';') {
                flushToken(word, tokens);
                throw QueryFailure.validation("Only one read-only SQL statement is allowed");
            }
            if (Character.isLetter(current) || current == '_') word.append(Character.toLowerCase(current));
            else flushToken(word, tokens);
        }
        flushToken(word, tokens);
        return tokens;
    }

    private void flushToken(StringBuilder word, List<String> tokens) {
        if (word.length() == 0) return;
        tokens.add(word.toString());
        word.setLength(0);
    }

    private boolean sameOrigin(URI base, URI target) {
        return java.util.Objects.equals(base.getScheme(), target.getScheme())
                && java.util.Objects.equals(base.getHost(), target.getHost())
                && effectivePort(base) == effectivePort(target);
    }

    private int effectivePort(URI uri) {
        if (uri.getPort() > 0) return uri.getPort();
        return uri.getScheme().equalsIgnoreCase("https") ? 443 : 80;
    }

    private int timeoutSeconds() {
        return Math.max(1, (int) Math.ceil(properties.timeout().toMillis() / 1000d));
    }

    private String nullToEmpty(String value) { return value == null ? "" : value; }

    private void checkSize(QueryTable table) {
        if (table.columns.size() > properties.maxColumns()) throw QueryFailure.validation("Query returned too many columns");
        if (table.rowCount > properties.maxRows()) throw QueryFailure.validation("Query returned too many rows");
        if (table.estimatedBytes > properties.maxResponseBytes() * 2L) throw QueryFailure.validation("Query response exceeds the configured byte limit");
    }

    /** Serialize only the final bounded result; intermediate steps use the columnar estimate. */
    private void checkFinalSize(QueryTable table) {
        try {
            if (mapper.writeValueAsBytes(Map.of("columns", table.columns, "rows", table.rows)).length > properties.maxResponseBytes()) {
                throw QueryFailure.validation("Query response exceeds the configured byte limit");
            }
        } catch (com.fasterxml.jackson.core.JsonProcessingException error) {
            throw QueryFailure.validation("Query response could not be serialized");
        }
    }

    /**
     * Immutable, bounded columnar table.  Narrowing and renaming reuse the
     * same chunks; only operators that inherently need global row ordering
     * materialize row views.
     */
    private static final class QueryTable {
        private static final int CHUNK_ROWS = 1_024;
        private final List<String> columns;
        private final List<QueryChunk> chunks;
        private final List<List<JsonNode>> rows;
        private final int rowCount;
        private final long estimatedBytes;

        private QueryTable(List<String> columns, List<List<JsonNode>> rows) {
            this(columns, chunkRows(columns, rows), true);
        }

        private QueryTable(List<String> columns, List<QueryChunk> chunks, boolean columnar) {
            this.columns = List.copyOf(columns);
            this.chunks = List.copyOf(chunks);
            this.rowCount = this.chunks.stream().mapToInt(QueryChunk::rowCount).sum();
            this.rows = new RowView(this.chunks, this.rowCount);
            this.estimatedBytes = estimateBytes(this.columns, this.chunks);
        }

        private static QueryTable fromChunks(List<String> columns, List<QueryChunk> chunks) {
            return new QueryTable(columns, chunks, true);
        }

        private QueryTable withColumns(List<String> nextColumns) {
            return fromChunks(nextColumns, chunks);
        }

        private static List<QueryChunk> chunkRows(List<String> columns, List<List<JsonNode>> rows) {
            List<QueryChunk> result = new ArrayList<>();
            List<List<JsonNode>> chunk = new ArrayList<>(CHUNK_ROWS);
            for (List<JsonNode> row : rows) {
                if (row.size() != columns.size()) throw QueryFailure.validation("Query row width is invalid");
                chunk.add(List.copyOf(row));
                if (chunk.size() == CHUNK_ROWS) {
                    result.add(QueryChunk.fromRows(chunk, columns.size()));
                    chunk = new ArrayList<>(CHUNK_ROWS);
                }
            }
            if (!chunk.isEmpty()) result.add(QueryChunk.fromRows(chunk, columns.size()));
            return result;
        }

        private static long estimateBytes(List<String> columns, List<QueryChunk> chunks) {
            long bytes = columns.stream().mapToLong(column -> column.length() * 2L + 8).sum();
            for (QueryChunk chunk : chunks) bytes += chunk.estimatedBytes();
            return bytes;
        }

        private int columnIndex(String name, String stepId) {
            int index = columns.indexOf(name);
            if (index < 0) throw QueryFailure.validation("Step " + stepId + " references missing column " + name);
            return index;
        }
    }

    private static final class QueryChunk {
        private final List<List<JsonNode>> columns;
        private final int rowCount;
        private final long estimatedBytes;

        private QueryChunk(List<List<JsonNode>> columns, int rowCount) {
            this.columns = columns.stream().map(List::copyOf).toList();
            this.rowCount = rowCount;
            long bytes = 0;
            for (List<JsonNode> column : this.columns) {
                for (JsonNode value : column) bytes += value == null || value.isNull() ? 4 : value.toString().length() * 2L + 4;
            }
            this.estimatedBytes = bytes;
        }

        private static QueryChunk fromRows(List<List<JsonNode>> rows, int columnCount) {
            List<List<JsonNode>> columns = new ArrayList<>(columnCount);
            for (int column = 0; column < columnCount; column++) {
                List<JsonNode> values = new ArrayList<>(rows.size());
                for (List<JsonNode> row : rows) values.add(row.get(column));
                columns.add(values);
            }
            return new QueryChunk(columns, rows.size());
        }

        private QueryChunk select(int[] indexes) {
            List<List<JsonNode>> selected = new ArrayList<>(indexes.length);
            for (int index : indexes) selected.add(columns.get(index));
            return new QueryChunk(selected, rowCount);
        }

        private int rowCount() { return rowCount; }
        private long estimatedBytes() { return estimatedBytes; }

        private List<JsonNode> row(int index) {
            List<JsonNode> row = new ArrayList<>(columns.size());
            for (List<JsonNode> column : columns) row.add(column.get(index));
            return List.copyOf(row);
        }
    }

    private static final class RowView extends java.util.AbstractList<List<JsonNode>> {
        private final List<QueryChunk> chunks;
        private final int[] starts;
        private final int size;

        private RowView(List<QueryChunk> chunks, int size) {
            this.chunks = chunks;
            this.size = size;
            this.starts = new int[chunks.size()];
            int offset = 0;
            for (int index = 0; index < chunks.size(); index++) {
                starts[index] = offset;
                offset += chunks.get(index).rowCount();
            }
        }

        @Override
        public List<JsonNode> get(int index) {
            if (index < 0 || index >= size) throw new IndexOutOfBoundsException(index);
            int low = 0;
            int high = starts.length - 1;
            while (low <= high) {
                int middle = (low + high) >>> 1;
                if (starts[middle] <= index) low = middle + 1;
                else high = middle - 1;
            }
            int chunkIndex = Math.max(0, high);
            return chunks.get(chunkIndex).row(index - starts[chunkIndex]);
        }

        @Override
        public int size() { return size; }
    }

    private static final class ActiveQuery {
        private final String actor;
        private final ExecutionControl control;
        private final Future<QueryTable> future;

        private ActiveQuery(String actor, ExecutionControl control, Future<QueryTable> future) {
            this.actor = actor;
            this.control = control;
            this.future = future;
        }

        private String actor() { return actor; }
        private ExecutionControl control() { return control; }
        private Future<QueryTable> future() { return future; }
    }

    /** Cancellation propagates into JDBC statements and HTTP request futures. */
    private static final class ExecutionControl {
        private final AtomicBoolean cancelled = new AtomicBoolean();
        private final AtomicReference<Statement> statement = new AtomicReference<>();
        private final AtomicReference<CompletableFuture<?>> request = new AtomicReference<>();
        private final AtomicReference<InputStream> body = new AtomicReference<>();

        private boolean cancelled() { return cancelled.get(); }

        private void ensureActive() {
            if (cancelled() || Thread.currentThread().isInterrupted()) throw QueryFailure.cancelled();
        }

        private void bind(Statement value) {
            ensureActive();
            statement.set(value);
            if (cancelled()) {
                try { value.cancel(); } catch (SQLException ignored) { }
                ensureActive();
            }
        }

        private void unbind(Statement value) { statement.compareAndSet(value, null); }

        private void bind(CompletableFuture<?> value) {
            ensureActive();
            request.set(value);
            if (cancelled()) {
                value.cancel(true);
                ensureActive();
            }
        }

        private void unbind(CompletableFuture<?> value) { request.compareAndSet(value, null); }

        private void bind(InputStream value) {
            ensureActive();
            body.set(value);
            if (cancelled()) {
                try { value.close(); } catch (IOException ignored) { }
                ensureActive();
            }
        }

        private void unbind(InputStream value) { body.compareAndSet(value, null); }

        private void cancel() {
            if (!cancelled.compareAndSet(false, true)) return;
            Statement currentStatement = statement.get();
            if (currentStatement != null) {
                try { currentStatement.cancel(); } catch (SQLException ignored) { }
            }
            CompletableFuture<?> currentRequest = request.get();
            if (currentRequest != null) currentRequest.cancel(true);
            InputStream currentBody = body.get();
            if (currentBody != null) {
                try { currentBody.close(); } catch (IOException ignored) { }
            }
        }
    }

    private record SortKey(String column, boolean ascending) {}
    private record Aggregate(String column, String function, String as) {}

    private static final class QueryFailure extends RuntimeException {
        private final ServiceException exception;
        private QueryFailure(ServiceException exception) { this.exception = exception; }
        static QueryFailure validation(String message) { return new QueryFailure(ServiceException.validation(message)); }
        static QueryFailure cancelled() { return new QueryFailure(ServiceException.timeout("Query was cancelled")); }
        String safeMessage() { return exception.getMessage(); }
        ServiceException exception() { return exception; }
    }
}
