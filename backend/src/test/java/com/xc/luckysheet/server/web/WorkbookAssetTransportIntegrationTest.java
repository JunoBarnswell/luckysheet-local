package com.xc.luckysheet.server.web;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.contract.CreateWorkbookRequest;
import com.xc.luckysheet.server.service.WorkbookAssetService;
import com.xc.luckysheet.server.service.WorkbookCatalogService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.test.context.TestPropertySource;
import org.springframework.test.web.servlet.MockMvc;

import java.security.MessageDigest;
import java.util.Base64;
import java.util.HexFormat;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.springframework.security.test.web.servlet.request.SecurityMockMvcRequestPostProcessors.jwt;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.put;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.jsonPath;

@SpringBootTest
@AutoConfigureMockMvc
@TestPropertySource(properties = {
        "DATABASE_URL=jdbc:h2:mem:asset_transport;DB_CLOSE_DELAY=-1", "DATABASE_USERNAME=sa", "DATABASE_PASSWORD=",
        "JPA_DDL_AUTO=validate", "FLYWAY_BASELINE_ON_MIGRATE=false", "luckysheet.auth.mode=oidc",
        "AUTH_ISSUER=https://issuer.test", "AUTH_AUDIENCE=test", "AUTH_JWKS_URL=https://issuer.test/.well-known/jwks.json",
        "COORDINATION_MULTI_INSTANCE=false", "COORDINATION_REDIS_ENABLED=false"
})
class WorkbookAssetTransportIntegrationTest {
    @Autowired private MockMvc http;
    @Autowired private WorkbookCatalogService catalog;
    @Autowired private WorkbookAssetService assets;
    @Autowired private com.xc.luckysheet.server.persistence.AssetEntityRepository repository;
    @Autowired private ObjectMapper mapper;

    @Test
    void binaryAssetTransportPersistsCanonicalMetadataAndRejectsInvalidRequestsBeforeWriting() throws Exception {
        String unitId = "asset-transport";
        catalog.create(new CreateWorkbookRequest(unitId, "Blocks", snapshot(unitId)), "owner");
        byte[] png = Base64.getDecoder().decode("iVBORw0KGgoAAAANSUhEUgAAADAAAAAgCAIAAADbtmxLAAAAN0lEQVR4nO3OQQ0AMAgEMFRgFNOTMRccjyYV0Op5p1R8ICQkJJQeCAkJCaUHQkJCQumBkJDQsg8dwKZ5fgcr3gAAAABJRU5ErkJggg==");
        String checksum = HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256").digest(png));
        String assetId = "asset-" + checksum;
        String url = "/api/workbooks/" + unitId + "/assets/" + assetId;
        var identity = new com.xc.luckysheet.server.persistence.AssetEntity.Id(unitId, assetId);
        http.perform(put(url).with(jwt().jwt(j -> j.subject("owner")))
                .contentType("image/png").header("X-Asset-Mime-Type", "image/png")
                .header("X-Content-SHA256", checksum).content(png))
                .andExpect(status().isUnsupportedMediaType()).andExpect(jsonPath("$.code").value("VALIDATION_ERROR"));
        assertFalse(repository.existsById(identity));
        http.perform(put(url).with(jwt().jwt(j -> j.subject("owner")))
                .contentType("application/octet-stream").header("X-Content-SHA256", checksum).content(png))
                .andExpect(status().isBadRequest()).andExpect(jsonPath("$.code").value("VALIDATION_ERROR"));
        assertFalse(repository.existsById(identity));
        http.perform(put(url).with(jwt().jwt(j -> j.subject("owner")))
                .contentType("application/octet-stream").header("X-Asset-Mime-Type", "image/png")
                .header("X-Content-SHA256", "0".repeat(64)).content(png))
                .andExpect(status().isBadRequest()).andExpect(jsonPath("$.code").value("VALIDATION_ERROR"));
        assertFalse(repository.existsById(identity));
        http.perform(put(url).with(jwt().jwt(j -> j.subject("outsider")))
                .contentType("application/octet-stream").header("X-Asset-Mime-Type", "image/png")
                .header("X-Content-SHA256", checksum).content(png))
                .andExpect(status().isForbidden()).andExpect(jsonPath("$.code").value("FORBIDDEN"));
        assertFalse(repository.existsById(identity));
        http.perform(put(url).with(jwt().jwt(j -> j.subject("owner")))
                .contentType("application/octet-stream").header("X-Asset-Mime-Type", "image/png")
                .header("X-Content-SHA256", checksum).content(png))
                .andExpect(status().isOk()).andExpect(jsonPath("$.schema").value("AssetRef"))
                .andExpect(jsonPath("$.assetId").value(assetId)).andExpect(jsonPath("$.mimeType").value("image/png"))
                .andExpect(jsonPath("$.width").doesNotExist()).andExpect(jsonPath("$.height").doesNotExist());
        assertArrayEquals(png, assets.get(unitId, assetId, "owner").getContent());
    }

    private com.fasterxml.jackson.databind.JsonNode snapshot(String unitId) throws Exception {
        return mapper.readTree("{\"schema\":\"WorkbookSnapshot\",\"version\":11,\"unitId\":\"" + unitId
                + "\",\"name\":\"Blocks\",\"dimensionMetrics\":{\"normalFontFamily\":\"Calibri\",\"normalFontSizePx\":14.6666666667,\"maximumDigitWidthPx\":7},\"calculationSettings\":{},\"editingOptions\":{\"allowEditDirectly\":true,\"moveAfterEnter\":true,\"enterDirection\":\"down\",\"formulaAutoComplete\":true,\"valueAutoComplete\":true,\"fixedDecimalPlaces\":null},\"dataModel\":{\"externalLinks\":[],\"sources\":[],\"tables\":[],\"relationships\":[],\"views\":[]},\"sheets\":[{\"kind\":\"worksheet\",\"id\":\"sheet-1\",\"name\":\"Sheet1\",\"rowCount\":1000,\"columnCount\":26,\"cells\":{},\"merges\":[],\"pane\":{\"kind\":\"none\"},\"defaultRowHeightPx\":20,\"defaultColumnWidthPx\":64,\"pivots\":[],\"sparklines\":[],\"drawings\":[],\"drawingPayloads\":{},\"hyperlinks\":[],\"review\":{\"notesByCell\":{},\"notesById\":{},\"threadIdsByCell\":{},\"threadsById\":{}}}]}");
    }
}
