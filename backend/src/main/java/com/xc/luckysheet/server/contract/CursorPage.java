package com.xc.luckysheet.server.contract;

import java.util.List;

/** Stable shape for bounded catalog and history reads. */
public record CursorPage<T>(List<T> items, String nextCursor) {
    public CursorPage {
        items = List.copyOf(items);
    }
}
