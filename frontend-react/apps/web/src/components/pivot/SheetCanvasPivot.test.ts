import assert from "node:assert/strict";
import { describe, it } from "node:test";
import type { CanvasSheetSnapshot } from "@react-sheets/spreadsheet-app";
import type { PivotGridProjection } from "@react-sheets/core-model";
import {
  findPivotProjectionCell,
  isPivotValueCell,
  pivotProjectionCellRenderData,
  resolvePivotProjectionHit,
} from "../SheetCanvas";

function projection(overrides: Partial<PivotGridProjection> = {}): PivotGridProjection {
  return {
    schema: "PivotGridProjection",
    pivotId: "pivot-1",
    sheetId: "sheet-1",
    target: { sheetId: "sheet-1", anchor: { row: 4, column: 3 } },
    occupiedRange: { sheetId: "sheet-1", startRow: 4, endRow: 5, startColumn: 3, endColumn: 4 },
    cells: [{ id: "pivot-1|r0|c0", pivotId: "pivot-1", row: 0, column: 0, kind: "value", value: 42, text: "42", sourceRowPaths: [{ sheetId: "sheet-1", row: 7 }] }],
    collision: { status: "clear", reasons: [], conflictingRanges: [] },
    refresh: { status: "ready", revision: 0, sourceRevision: "source-1" },
    ...overrides,
  };
}

function sheet(pivot: PivotGridProjection): CanvasSheetSnapshot {
  return {
    id: "sheet-1",
    name: "Sheet1",
    columns: ["A", "B", "C", "D", "E"],
    columnCount: 5,
    rowCount: 20,
    occupiedCellCount: 0,
    getCell: () => undefined,
    usedRange: { sheetId: "sheet-1", startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
    drawings: [],
    drawingPayloads: new Map(),
    pivots: [],
    pivotResults: {},
    pivotProjections: { [pivot.pivotId]: pivot },
    sparklines: [],
    conditionalFormats: [],
    dataValidations: [],
    merges: [],
    pane: { kind: 'none' },
    defaultRowHeightPx: 20,
    defaultColumnWidthPx: 64,
    maximumDigitWidthPx: 7,
    rowHeightsPx: {},
    columnWidthsPx: {},
    hiddenRows: [],
    hiddenColumns: [],
    outlineGroups: [],
    outlineControls: [],
    filterRangeColumns: [],
    activeFilterColumns: [],
    filterButtons: [],
    filterButtonStates: [],
    getFilterColorDomain: () => [],
    getFilterIconDomain: () => [],
    getFilterValueDomain: () => [],
    getFilterCriterion: () => undefined,
    getFilterOwner: () => undefined,
    getActiveAutoFilter: () => undefined,
    sheetTables: [],
    previewRows: [],
  };
}

describe("SheetCanvas Pivot projection boundary", () => {
  it("maps anchor-relative projection cells to absolute worksheet coordinates", () => {
    const target = findPivotProjectionCell(sheet(projection()), 4, 3);
    assert.equal(target?.cell.id, "pivot-1|r0|c0");
    assert.equal(findPivotProjectionCell(sheet(projection()), 0, 0), null);
  });

  it("does not hit or render a colliding projection", () => {
    const colliding = projection({ collision: { status: "collision", reasons: ["cell-data"], conflictingRanges: [] } });
    const current = sheet(colliding);
    assert.equal(findPivotProjectionCell(current, 4, 3), null);
    assert.equal(resolvePivotProjectionHit(current, 4, 3), null);
  });

  it("returns a resolved Pivot context and keeps value details metadata", () => {
    const current = sheet(projection());
    const hit = resolvePivotProjectionHit(current, 4, 3);
    assert.equal(hit?.kind, "pivot");
    assert.equal(hit?.objectId, "pivot-1");
    assert.equal(hit?.pivot?.row, 4);
    const cell = findPivotProjectionCell(current, 4, 3)?.cell;
    assert.ok(cell);
    assert.equal(isPivotValueCell(cell), true);
    assert.equal(pivotProjectionCellRenderData(cell).displayValue, "42");
  });
});
