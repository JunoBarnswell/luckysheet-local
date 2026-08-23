package com.xc.luckysheet.server.mutation;

/**
 * Server-owned concurrency policy.  Exact-base is deliberately the only
 * accepted rule until a mutation has a reducer whose transform is proven
 * deterministic for both the Java snapshot and the browser model.
 */
public enum MutationRebasePolicy {
    EXACT_BASE
}
