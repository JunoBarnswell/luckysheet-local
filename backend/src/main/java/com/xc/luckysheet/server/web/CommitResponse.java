package com.xc.luckysheet.server.web;
import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;
import com.fasterxml.jackson.databind.JsonNode;
public record CommitResponse(CommittedOperationEnvelope operation, JsonNode changeSet) { }
