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
        var snapshot = mapper.readTree("{\"unitId\":\"book-context\",\"name\":\"Context\",\"sheets\":[]}");
        catalog.create(new CreateWorkbookRequest("book-context", "Context", snapshot), "actor-context");
        var summaries = catalog.list("actor-context", "recent", null, null, null);
        org.junit.jupiter.api.Assertions.assertEquals(1, summaries.size());
        org.junit.jupiter.api.Assertions.assertEquals("owner", summaries.get(0).role().wireValue());
        org.junit.jupiter.api.Assertions.assertEquals(java.util.List.of("我的云文档"), summaries.get(0).locationPath());
        var state = catalog.putUserState("book-context", new UserStateRequest(true, null), "actor-context");
        org.junit.jupiter.api.Assertions.assertTrue(state.favorite());
        org.junit.jupiter.api.Assertions.assertEquals(1, workspace.list("actor-context").size());
    }
}
