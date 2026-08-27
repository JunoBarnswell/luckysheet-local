import assert from "node:assert/strict";
import { describe, it } from "node:test";
import type { CanvasSheetSnapshot } from "@react-sheets/spreadsheet-app";
import type { PivotGridProjection, PivotReportFilterSummary } from "@react-sheets/core-model";
import {
  findPivotProjectionCell,
  isPivotValueCell,
  pivotFilterSummaryText,
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
    collision: { status: "clear", reasons: [], conflictingRanges: [], conflicts: [] },
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
    dataRegions: [],
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
    getFilterDomainDescriptor: () => ({ column: 0, values: [], scalarTypes: [], dominantType: 'empty', hasBlank: false, dateDomain: [], dateHierarchy: [], colorDomain: [], iconDomain: [], supportedFamilies: ['values'] }),
    getFilterCriterion: () => undefined,
    getFilterOwner: () => undefined,
    getActiveAutoFilter: () => undefined,
    sheetTables: [],
    forEachOccupiedCell: () => {},
  };
}

describe("SheetCanvas Pivot projection boundary", () => {
  it("maps anchor-relative projection cells to absolute worksheet coordinates", () => {
    const target = findPivotProjectionCell(sheet(projection()), 4, 3);
    assert.equal(target?.cell.id, "pivot-1|r0|c0");
    assert.equal(findPivotProjectionCell(sheet(projection()), 0, 0), null);
  });

  it("does not hit or render a colliding projection", () => {
    const colliding = projection({ collision: { status: "collision", reasons: ["cell-data"], conflictingRanges: [], conflicts: [] } });
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

  it("localizes semantic Pivot captions without changing the projection model", () => {
    const cell = { id: 'pivot-1|caption', pivotId: 'pivot-1', row: 0, column: 0, kind: 'column-header' as const, value: null, text: 'Row Labels', captionKey: 'row-labels' as const };
    assert.equal(pivotProjectionCellRenderData(cell, 'zh-CN').displayValue, '行标签');
    assert.equal(pivotProjectionCellRenderData(cell, 'en-US').displayValue, 'Row Labels');
  });

  it('renders localized, family-aware report summaries and keeps All only for inactive filters', () => {
    const all: PivotReportFilterSummary = { fieldName: 'Region', active: false, entries: [{ kind: 'manual', family: 'manual', active: false, mode: 'all', count: 0, memberValues: [] }] };
    assert.equal(pivotFilterSummaryText(all, 'en-US'), 'Region: All');
    const active: PivotReportFilterSummary = {
      fieldName: 'Region',
      active: true,
      entries: [
        { kind: 'manual', family: 'manual', active: false, mode: 'all', count: 0, memberValues: [] },
        { kind: 'condition', family: 'label', active: true, operator: 'begins-with', value: 'N' },
        { kind: 'condition', family: 'date', active: true, operator: 'between', value: '2024-01-01', value2: '2024-12-31' },
        { kind: 'top-items', family: 'top-items', active: true, direction: 'top', mode: 'items', threshold: 3, valueFieldName: 'Amount' },
      ],
    };
    const text = pivotFilterSummaryText(active, 'en-US');
    assert.match(text, /Region:/);
    assert.match(text, /Begins With/);
    assert.match(text, /Between/);
    assert.match(text, /Top Items 3 by Amount/);
    assert.doesNotMatch(text, /Region: All/);
    assert.match(pivotFilterSummaryText(active, 'zh-CN'), /标签筛选/);
  });

  it('resolves Pivot presentation style and options instead of using one fixed palette', () => {
    const cell = { id: 'pivot-1|header', pivotId: 'pivot-1', row: 1, column: 1, kind: 'column-header' as const, value: null, text: 'Amount' };
    const medium = pivotProjectionCellRenderData(cell, 'en-US', {
      styleName: 'PivotStyleMedium4',
      styleOptions: { showRowHeaders: true, showColumnHeaders: true, showRowStripes: false, showColumnStripes: false, showLastColumn: false },
    });
    assert.equal(medium.style?.background, '#d9e2f3');

    const striped = pivotProjectionCellRenderData({ ...cell, kind: 'value', row: 2, column: 2, value: 10, text: '10', isLastColumn: true }, 'en-US', {
      styleName: 'PivotStyleLight16',
      styleOptions: { showRowHeaders: false, showColumnHeaders: false, showRowStripes: true, showColumnStripes: false, showLastColumn: true },
    });
    assert.equal(striped.style?.background, '#e0ecff');
    assert.equal(striped.style?.bold, false);
  });

  it('renders Pivot empty and error values through display options', () => {
    const empty = { id: 'pivot-1|empty', pivotId: 'pivot-1', row: 1, column: 1, kind: 'value' as const, value: null, text: '—' };
    const error = { id: 'pivot-1|error', pivotId: 'pivot-1', row: 1, column: 2, kind: 'value' as const, value: { kind: 'error' as const, code: '#N/A' as const }, text: 'ERR' };
    const presentation = {
      styleOptions: { showRowHeaders: true, showColumnHeaders: true, showRowStripes: false, showColumnStripes: false, showLastColumn: false },
      displayOptions: { fillEmptyCells: true, emptyCellText: '—', showErrorValues: true, errorCellText: 'ERR', showFieldHeaders: true, autoFitColumnsOnUpdate: true },
    };
    assert.equal(pivotProjectionCellRenderData(empty, 'en-US', presentation).displayValue, '—');
    const errorData = pivotProjectionCellRenderData(error, 'en-US', presentation);
    assert.equal(errorData.displayValue, 'ERR');
    assert.equal(errorData.invalid, true);
  });
});
