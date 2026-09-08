package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.AnalyticsExecutionRequest;
import com.xc.luckysheet.server.contract.AnalyticsExecutionResponse;
import com.xc.luckysheet.server.contract.AnalyticsPrepareRequest;
import com.xc.luckysheet.server.contract.AnalyticsPrepareResponse;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.QueryExecutionService;
import jakarta.validation.Valid;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ObjectNode;

/** Native analytics transport. Workbook mutations remain on WorkbookController. */
@RestController
@RequestMapping("/api/workbooks/{unitId}/analytics")
public class AnalyticsController {
    private final QueryExecutionService analytics;

    public AnalyticsController(QueryExecutionService analytics) {
        this.analytics = analytics;
    }

    @PostMapping("/prepare")
    public AnalyticsPrepareResponse prepare(@PathVariable String unitId,
            @Valid @RequestBody AnalyticsPrepareRequest request, Authentication authentication) {
        return analytics.prepareAnalytics(unitId, request, actor(authentication));
    }

    @PostMapping("/{queryId}/execute")
    public JsonNode execute(@PathVariable String unitId, @PathVariable String queryId,
            @Valid @RequestBody AnalyticsExecutionRequest request, Authentication authentication) {
        return wire(analytics.executeAnalytics(unitId, queryId, request, actor(authentication)));
    }

    @PostMapping("/{queryId}/viewport")
    public JsonNode viewport(@PathVariable String unitId, @PathVariable String queryId,
            @Valid @RequestBody AnalyticsExecutionRequest request, Authentication authentication) {
        return wire(analytics.viewportAnalytics(unitId, queryId, request, actor(authentication)));
    }

    @PostMapping("/{queryId}/drilldown")
    public JsonNode drilldown(@PathVariable String unitId, @PathVariable String queryId,
            @Valid @RequestBody AnalyticsExecutionRequest request, Authentication authentication) {
        return wire(analytics.drilldownAnalytics(unitId, queryId, request, actor(authentication)));
    }

    @PostMapping("/{queryId}/cancel")
    public ResponseEntity<Void> cancel(@PathVariable String unitId, @PathVariable String queryId,
            Authentication authentication) {
        analytics.cancelAnalytics(unitId, queryId, actor(authentication));
        return ResponseEntity.accepted().build();
    }

    @DeleteMapping("/{queryId}")
    public ResponseEntity<Void> delete(@PathVariable String unitId, @PathVariable String queryId,
            Authentication authentication) {
        analytics.cancelAnalytics(unitId, queryId, actor(authentication));
        return ResponseEntity.accepted().build();
    }

    private static String actor(Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return ActorIdentity.subject(authentication);
    }

    /** Keep the native AnalyticsResponse at the wire root; proof metadata is
     * additive and consumers can pass the body directly to the Rust contract. */
    private static ObjectNode wire(AnalyticsExecutionResponse response) {
        ObjectNode body = (ObjectNode) response.result().deepCopy();
        body.put("queryId", response.queryId());
        body.put("sourceRevision", response.sourceRevision());
        body.put("executionToken", response.executionToken());
        body.put("resultHash", response.resultHash());
        body.put("executedAt", response.executedAt().toString());
        body.put("durationMs", response.durationMs());
        return body;
    }
}
