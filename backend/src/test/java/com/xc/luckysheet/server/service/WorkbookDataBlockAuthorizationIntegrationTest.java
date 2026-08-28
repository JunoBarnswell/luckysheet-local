package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.contract.ShareCreateRequest;
import com.xc.luckysheet.server.contract.WorkbookAclRole;
import com.xc.luckysheet.server.store.DataBlockRow;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.TestPropertySource;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.Instant;
import java.util.HexFormat;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@SpringBootTest
@TestPropertySource(properties = {
        "DATABASE_URL=jdbc:h2:mem:block_authorization;DB_CLOSE_DELAY=-1",
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
class WorkbookDataBlockAuthorizationIntegrationTest {
    @Autowired private WorkbookCatalogService catalog;
    @Autowired private GuestShareService shares;
    @Autowired private WorkbookDataBlockCommitService commits;
    @Autowired private WorkbookDataBlockStore blocks;
    @Autowired private ObjectMapper mapper;

    @Test
    void revokedEditorCannotCommitBytesReadBeforeTheWriteBoundary() throws Exception {
        String unitId = "block-revoked";
        String owner = "owner-revoked";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), owner);
        var share = shares.create(unitId, new ShareCreateRequest("editor", Instant.now().plusSeconds(600)), owner);
        byte[] content = "blocked-after-revocation".getBytes(StandardCharsets.UTF_8);

        shares.revoke(unitId, share.shareId(), owner);
        ServiceException error = assertThrows(ServiceException.class, () -> commits.commit(
                row(unitId, content), WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_BYTES,
                WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_COUNT, "guest:" + share.shareId()));

        assertEquals("FORBIDDEN", error.code());
        assertTrue(blocks.find(unitId, "source", "block").isEmpty());
    }

    @Test
    void editorCommitThatCompletesBeforeRevocationRemainsPersisted() throws Exception {
        String unitId = "block-committed";
        String owner = "owner-committed";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), owner);
        var share = shares.create(unitId, new ShareCreateRequest("editor", Instant.now().plusSeconds(600)), owner);
        byte[] content = "committed-before-revocation".getBytes(StandardCharsets.UTF_8);

        var metadata = commits.commit(row(unitId, content), WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_BYTES,
                WorkbookDataBlockService.MAX_WORKBOOK_BLOCK_COUNT, "guest:" + share.shareId());
        shares.revoke(unitId, share.shareId(), owner);

        assertEquals(content.length, metadata.byteLength());
        assertEquals(HexFormat.of().formatHex(content), HexFormat.of().formatHex(blocks.find(unitId, "source", "block").orElseThrow().content()));
    }

    private DataBlockRow row(String unitId, byte[] content) throws Exception {
        return new DataBlockRow(unitId, "source", "block", HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(content)),
                content.length, content, Instant.now(), Instant.now());
    }

    private com.fasterxml.jackson.databind.JsonNode snapshot(String unitId) throws Exception {
        return mapper.readTree("{\"schema\":\"WorkbookSnapshot\",\"version\":9,\"unitId\":\"" + unitId
                + "\",\"name\":\"Blocks\",\"dimensionMetrics\":{\"normalFontFamily\":\"Calibri\",\"normalFontSizePx\":14.6666666667,\"maximumDigitWidthPx\":7},\"calculationSettings\":{},\"editingOptions\":{\"allowEditDirectly\":true,\"moveAfterEnter\":true,\"enterDirection\":\"down\",\"formulaAutoComplete\":true,\"valueAutoComplete\":true,\"fixedDecimalPlaces\":null},\"dataModel\":{\"sources\":[],\"tables\":[],\"relationships\":[],\"views\":[]},\"sheets\":[{\"kind\":\"worksheet\",\"id\":\"sheet-1\",\"name\":\"Sheet1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{},\"merges\":[],\"pane\":{\"kind\":\"none\"},\"defaultRowHeightPx\":20,\"defaultColumnWidthPx\":64,\"pivots\":[],\"sparklines\":[],\"drawings\":[],\"drawingPayloads\":{},\"review\":{\"notesByCell\":{},\"notesById\":{},\"threadIdsByCell\":{},\"threadsById\":{}}}]}");
    }
}
