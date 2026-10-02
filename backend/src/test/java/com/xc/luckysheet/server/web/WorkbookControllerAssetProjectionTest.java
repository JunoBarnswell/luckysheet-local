package com.xc.luckysheet.server.web;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class WorkbookControllerAssetProjectionTest {
    private final ObjectMapper mapper = new ObjectMapper();

    @Test
    void servesOnlyAssetsReferencedByTheSubjectSafeSnapshot() throws Exception {
        var visibleSnapshot = mapper.readTree("""
                {"sheets":[{"cells":{"0":{"0":{"presentation":{"asset":{"schema":"AssetRef","assetId":"asset-visible"}}}}},
                  "drawings":[{"payloadId":"drawing-1"}],
                  "drawingPayloads":{"drawing-1":{"asset":{"schema":"AssetRef","assetId":"asset-drawing"}}}}]}
                """);
        assertTrue(WorkbookController.snapshotReferencesAsset(visibleSnapshot, "asset-visible"));
        assertTrue(WorkbookController.snapshotReferencesAsset(visibleSnapshot, "asset-drawing"));
    }

    @Test
    void doesNotServeAssetWhenProjectionRemovedItsHiddenCellReference() throws Exception {
        var projectedSnapshot = mapper.readTree("""
                {"sheets":[{"cells":{"0":{"0":{"value":"visible"}}},"drawings":[],"drawingPayloads":{}}]}
                """);
        assertFalse(WorkbookController.snapshotReferencesAsset(projectedSnapshot, "asset-hidden"));
    }
}
