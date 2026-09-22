package com.xc.luckysheet.server.service;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonToken;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.xc.luckysheet.server.store.WorkbookStore;
import org.springframework.stereotype.Service;

import java.io.IOException;

/** Protects canonical block references, including retained restore and undo history. */
@Service
public class WorkbookDataBlockReferenceGuard {
    private final WorkbookStore store;
    private final ObjectMapper mapper;

    public WorkbookDataBlockReferenceGuard(WorkbookStore store, ObjectMapper mapper) {
        this.store = store;
        this.mapper = mapper;
    }

    /** Must run inside the same workbook write transaction as the subsequent delete. */
    public void requireUnreferenced(String unitId, String sourceId, String blockId) {
        var workbook = store.find(unitId).orElseThrow(() -> ServiceException.notFound("Workbook not found"));
        inspect(workbook.snapshotJson(), sourceId, blockId);
        store.forEachRetainedHistoryDocument(unitId, document -> inspect(document, sourceId, blockId));
    }

    private void inspect(String document, String sourceId, String blockId) {
        try {
            if (document == null) throw new IOException("Missing retained document");
            try (JsonParser parser = mapper.getFactory().createParser(document)) {
                if (parser.nextToken() != JsonToken.START_OBJECT) throw new IOException("Expected a retained JSON object");
                if (containsReference(parser, sourceId, blockId)) {
                    throw new ServiceException("DATA_BLOCK_REFERENCED", 409,
                            "Data block " + sourceId + "/" + blockId + " is referenced by workbook state or retained history; only unreferenced staging blocks may be deleted");
                }
                if (parser.nextToken() != null) throw new IOException("Unexpected trailing JSON content");
            }
        } catch (IOException error) {
            throw new ServiceException("DATA_BLOCK_REFERENCE_CHECK_FAILED", 409,
                    "Cannot check references for data block " + sourceId + "/" + blockId + "; repair the retained workbook document before deleting", error);
        }
    }

    /** Streaming traversal preserves object boundaries and does not materialize cell grids. */
    private boolean containsReference(JsonParser parser, String sourceId, String blockId) throws IOException {
        if (parser.currentToken() == JsonToken.START_OBJECT) {
            String id = null;
            String dataSourceId = null;
            while (parser.nextToken() != JsonToken.END_OBJECT) {
                if (parser.currentToken() != JsonToken.FIELD_NAME) throw new IOException("Invalid retained object");
                String field = parser.currentName();
                if (parser.nextToken() == null) throw new IOException("Missing retained value");
                if (parser.currentToken() == JsonToken.VALUE_STRING) {
                    if ("id".equals(field)) id = parser.getText();
                    if ("dataSourceId".equals(field)) dataSourceId = parser.getText();
                } else if (containsReference(parser, sourceId, blockId)) {
                    return true;
                }
            }
            return blockId.equals(id) && sourceId.equals(dataSourceId);
        }
        if (parser.currentToken() == JsonToken.START_ARRAY) {
            while (parser.nextToken() != JsonToken.END_ARRAY) {
                if (parser.currentToken() == null) throw new IOException("Unclosed retained array");
                if (containsReference(parser, sourceId, blockId)) return true;
            }
        }
        return false;
    }
}
