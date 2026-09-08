package com.xc.luckysheet.server.contract;
import com.fasterxml.jackson.databind.JsonNode;
public record WorkbookImportResponse(String unitId, long revision, String checksum, WorkbookSummary summary,
                                     JsonNode manifest, WorkbookArtifactResponse artifact) { }
