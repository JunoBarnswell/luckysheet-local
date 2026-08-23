package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.databind.JsonNode;
import jakarta.validation.Valid;

import java.util.List;

/** Sanitized query intent. Source credentials are resolved only by sourceRef on the server. */
public record QueryExecutionRequest(
        @JsonProperty("queryId") String queryId,
        @JsonProperty("name") String name,
        @JsonProperty("connectorId") String connectorId,
        @JsonProperty("sourceRef") String sourceRef,
        @JsonProperty("statement") String statement,
        @JsonProperty("method") String method,
        @JsonProperty("body") JsonNode body,
        @JsonProperty("parameters") List<JsonNode> parameters,
        @JsonProperty("steps") @Valid List<QueryStep> steps
) {
    @JsonCreator
    public QueryExecutionRequest {
        if (queryId == null || queryId.isBlank() || queryId.length() > 200) throw new IllegalArgumentException("queryId is required");
        if (name == null || name.isBlank() || name.length() > 500) throw new IllegalArgumentException("query name is required");
        if (connectorId == null || connectorId.isBlank() || connectorId.length() > 50) throw new IllegalArgumentException("connectorId is required");
        if (sourceRef == null || sourceRef.isBlank() || sourceRef.length() > 200) throw new IllegalArgumentException("sourceRef is required");
        if (statement == null || statement.isBlank() || statement.length() > 100_000) throw new IllegalArgumentException("query statement is required");
        if (method != null && !method.isBlank() && !method.equalsIgnoreCase("GET") && !method.equalsIgnoreCase("POST")) {
            throw new IllegalArgumentException("REST query method must be GET or POST");
        }
        steps = steps == null ? List.of() : List.copyOf(steps);
        if (steps.size() > 100) throw new IllegalArgumentException("A query may contain at most 100 steps");
        parameters = parameters == null ? List.<JsonNode>of() : parameters.stream().<JsonNode>map(value -> value == null ? null : value.deepCopy()).toList();
        if (parameters.size() > 1000) throw new IllegalArgumentException("A query may contain at most 1000 parameters");
        body = body == null ? null : body.deepCopy();
    }
}
