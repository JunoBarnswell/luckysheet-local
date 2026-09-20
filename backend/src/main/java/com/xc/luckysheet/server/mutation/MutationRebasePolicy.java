package com.xc.luckysheet.server.mutation;

/**
 * Server-owned concurrency policy. Disjoint cell writes retain their exact
 * address and are allowed only when no intervening structural change exists.
 */
public enum MutationRebasePolicy {
    EXACT_BASE,
    DISJOINT_CELLS
}
