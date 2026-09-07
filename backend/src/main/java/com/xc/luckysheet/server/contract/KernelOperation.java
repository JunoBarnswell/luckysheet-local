package com.xc.luckysheet.server.contract;

/** Canonical operations accepted by workbook-kernel-host protocol v1. */
public enum KernelOperation {
    INIT("init"),
    OPEN("open"),
    CREATE("create"),
    MANIFEST("manifest"),
    CELL_GET("cell.get"),
    RANGE_GET("range.get"),
    PAGE_GET("page.get"),
    PAGE_LOAD("page.load"),
    COMMAND("command"),
    CLOSE("close"),
    FORMULA_EVALUATE("formula.evaluate"),
    FORMULA_RECALCULATE("formula.recalculate"),
    ANALYTICS_EXECUTE("analytics.execute"),
    GEOMETRY_COMPUTE_PANE_MAP("geometry.computePaneMap"),
    GEOMETRY_HIT_TEST("geometry.hitTest"),
    GEOMETRY_CELL_RECT("geometry.cellRect"),
    GEOMETRY_HEADER_RECT("geometry.headerRect"),
    DOCUMENT_IMPORT("document.import"),
    DOCUMENT_EXPORT("document.export");

    private final String wireName;

    KernelOperation(String wireName) { this.wireName = wireName; }

    public String wireName() { return wireName; }
}
