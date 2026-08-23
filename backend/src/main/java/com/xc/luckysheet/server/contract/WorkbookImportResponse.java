package com.xc.luckysheet.server.contract;

import com.fasterxml.jackson.databind.JsonNode;

public record WorkbookImportResponse(
        String unitId,
        long revision,
        String checksum,
        WorkbookSummary summary,
        JsonNode snapshot,
        WorkbookArtifactResponse artifact
) {
    public WorkbookImportResponse(WorkbookSummary summary, JsonNode snapshot, WorkbookArtifactResponse artifact) {
        this(summary.unitId(), summary.revision(), null, summary, snapshot, artifact);
    }
}
