package com.xc.luckysheet.server.contract;
import com.fasterxml.jackson.databind.JsonNode;
/** Revision-pinned manifest; cells are exclusively delivered through the page API. */
public record WorkbookOpenResponse(String unitId, long revision, JsonNode manifest, String checksum) { }
