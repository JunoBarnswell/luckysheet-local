package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.CommittedOperationEnvelope;

public record CommitResponse(CommittedOperationEnvelope operation) {
}
