package com.xc.luckysheet.server.contract;

import jakarta.validation.constraints.NotNull;

import java.util.List;

public record AssetReconcileRequest(@NotNull List<String> assetIds) {
}
