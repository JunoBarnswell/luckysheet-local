package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.xc.luckysheet.server.store.WorkbookDataBlockStore;
import org.springframework.stereotype.Service;

import java.util.HashMap;
import java.util.Map;

/** Validates references before publication while the caller owns the workbook write lock. */
@Service
public class WorkbookDataBlockPublicationGuard {
    private final WorkbookDataBlockStore blocks;

    public WorkbookDataBlockPublicationGuard(WorkbookDataBlockStore blocks) {
        this.blocks = blocks;
    }

    public void requireNewReferences(String unitId, JsonNode previous, JsonNode candidate) {
        Map<String, JsonNode> previousBlocks = new HashMap<>();
        for (JsonNode source : sources(previous)) previousBlocks.put(source.path("id").asText(), source.path("blocks"));
        requireReferences(unitId, candidate, previousBlocks);
    }

    /** Restoring a historical snapshot rechecks all referenced metadata, including previously missing bytes. */
    public void requireSnapshot(String unitId, JsonNode candidate) {
        requireReferences(unitId, candidate, Map.of());
    }

    /**
     * Copies only the immutable blocks published by the source snapshot. The
     * source metadata is verified first, and the database performs the byte
     * copy without loading payloads into the application heap.
     */
    public void copySnapshotReferences(String sourceUnitId, String targetUnitId, JsonNode candidate) {
        requireSnapshot(sourceUnitId, candidate);
        for (JsonNode source : sources(candidate)) {
            String sourceId = source.path("id").asText();
            for (JsonNode ref : source.path("blocks")) {
                String blockId = ref.path("id").asText();
                if (blocks.copyToWorkbook(sourceUnitId, targetUnitId, sourceId, blockId) != 1) {
                    throw new ServiceException("DATA_BLOCK_MISSING", 409,
                            "Data block " + sourceId + "/" + blockId + " disappeared while copying the workbook; retry from a consistent source");
                }
            }
        }
    }

    private void requireReferences(String unitId, JsonNode candidate, Map<String, JsonNode> previousBlocks) {
        for (JsonNode source : sources(candidate)) {
            String sourceId = source.path("id").asText();
            JsonNode refs = source.path("blocks");
            if (sourceId.isBlank() || !refs.isArray()) throw ServiceException.validation("Invalid data source block references");
            if (refs.equals(previousBlocks.get(sourceId))) continue;
            for (JsonNode ref : refs) {
                String blockId = ref.path("id").asText();
                if (blockId.isBlank() || !sourceId.equals(ref.path("dataSourceId").asText())) {
                    throw ServiceException.validation("Data block reference identity does not match its source");
                }
                var metadata = blocks.findMetadata(unitId, sourceId, blockId).orElseThrow(() -> new ServiceException(
                        "DATA_BLOCK_MISSING", 409, "Data block " + sourceId + "/" + blockId + " is not uploaded in this workbook; upload it before publishing the reference"));
                if (!ref.path("byteLength").isIntegralNumber() || ref.path("byteLength").longValue() != metadata.byteLength()
                        || !metadata.checksum().equals(ref.path("checksum").asText())) {
                    throw new ServiceException("DATA_BLOCK_METADATA_MISMATCH", 409,
                            "Data block " + sourceId + "/" + blockId + " does not match its uploaded checksum or length; publish the matching descriptor or upload a new block id");
                }
            }
        }
    }

    private JsonNode sources(JsonNode snapshot) {
        JsonNode sources = snapshot.path("dataModel").path("sources");
        if (!sources.isArray()) throw ServiceException.validation("Canonical data source collection is required before publishing blocks");
        return sources;
    }
}
