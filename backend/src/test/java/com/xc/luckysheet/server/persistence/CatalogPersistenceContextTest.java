package com.xc.luckysheet.server.persistence;

import com.xc.luckysheet.server.service.WorkbookCatalogService;
import com.xc.luckysheet.server.service.WorkspaceService;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.UserStateRequest;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.TestPropertySource;

import static org.junit.jupiter.api.Assertions.assertNotNull;

@SpringBootTest
@TestPropertySource(properties = {
        "DATABASE_URL=jdbc:h2:mem:catalog_context;DB_CLOSE_DELAY=-1",
        "DATABASE_USERNAME=sa",
        "DATABASE_PASSWORD=",
        "JPA_DDL_AUTO=validate",
        "FLYWAY_BASELINE_ON_MIGRATE=false",
        "AUTH_ISSUER=https://issuer.test",
        "AUTH_AUDIENCE=test",
        "AUTH_JWKS_URL=https://issuer.test/.well-known/jwks.json",
        "COORDINATION_MULTI_INSTANCE=false",
        "COORDINATION_REDIS_ENABLED=false"
})
class CatalogPersistenceContextTest {
    @Autowired
    private WorkbookCatalogService catalog;

    @Autowired
    private WorkbookSourceArtifactEntityRepository artifacts;

    @Autowired
    private WorkspaceService workspace;

    @Autowired
    private ObjectMapper mapper;

    @Test
    void flywayCreatesCatalogSchemaAndSpringValidatesTheEntityModel() {
        assertNotNull(catalog);
        assertNotNull(artifacts);
    }

    @Test
    void catalogCreatesPersonalSpaceAndReturnsOneActorEnrichedSummary() throws Exception {
        var snapshot = mapper.readTree("""
                {"schema":"WorkbookSnapshot","version":5,"unitId":"book-context","name":"Context","dimensionMetrics":{"normalFontFamily":"Calibri","normalFontSizePx":14.6666666667,"maximumDigitWidthPx":7},"dataModel":{"sources":[],"tables":[],"relationships":[],"views":[]},"sheets":[{"kind":"worksheet","id":"sheet-1","name":"Sheet1","rowCount":1000,"columnCount":26,"cells":{},"merges":[],"pane":{"kind":"none"},"defaultRowHeightPx":20,"defaultColumnWidthPx":64,"pivots":[],"sparklines":[],"drawings":[],"drawingPayloads":{}}]}
                """);
        catalog.create(new CreateWorkbookRequest("book-context", "Context", snapshot), "actor-context");
        var summaries = catalog.list("actor-context", "recent", null, null, null, 0, 50);
        org.junit.jupiter.api.Assertions.assertEquals(1, summaries.items().size());
        org.junit.jupiter.api.Assertions.assertEquals("owner", summaries.items().get(0).role().wireValue());
        org.junit.jupiter.api.Assertions.assertEquals(java.util.List.of("我的云文档"), summaries.items().get(0).locationPath());
        var state = catalog.putUserState("book-context", new UserStateRequest(true, null), "actor-context");
        org.junit.jupiter.api.Assertions.assertTrue(state.favorite());
        org.junit.jupiter.api.Assertions.assertEquals(1, workspace.list("actor-context").size());
    }
}
