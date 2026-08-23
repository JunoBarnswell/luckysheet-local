import type { WorkbookSnapshot, WorksheetPane } from '@react-sheets/core-model';
import { verifyXlsxSourceArtifact, xlsxArtifactNeedsLayoutRepair } from './source-artifact';
import { importXlsx } from './import';
import type { XlsxSourceArtifact } from './types';

export interface XlsxSheetLayoutRepair {
  sheetId: string;
  sourceSheetName: string;
  defaultRowHeightPx: number;
  defaultColumnWidthPx: number;
  rowHeightsPx: Record<number, number>;
  columnWidthsPx: Record<number, number>;
  pane: WorksheetPane;
  fontSizesPx: Array<{ row: number; column: number; fontSizePx: number }>;
}

export interface XlsxLayoutRepairPlan {
  schema: 'XlsxLayoutRepairPlan';
  sourceChecksum: string;
  sheets: XlsxSheetLayoutRepair[];
  summary: { sheets: number; defaults: number; rows: number; columns: number; fonts: number; panes: number };
  report: NonNullable<XlsxSourceArtifact['capabilityReport']>;
}

export async function createXlsxLayoutRepairPlan(snapshot: WorkbookSnapshot, artifact: XlsxSourceArtifact): Promise<XlsxLayoutRepairPlan> {
  await verifyXlsxSourceArtifact(artifact);
  if (!xlsxArtifactNeedsLayoutRepair(artifact)) throw new Error('This workbook was imported with the current XLSX codec and does not need layout repair');
  const imported = await importXlsx({ fileName: artifact.fileName, buffer: artifact.buffer, options: { compatibilityTarget: 'B', compatibilityMode: 'balanced' } });
  const parsed = imported.snapshot;
  const sheets: XlsxSheetLayoutRepair[] = [];
  let defaults = 0;
  let rows = 0;
  let columns = 0;
  let fonts = 0;
  let panes = 0;
  for (const current of snapshot.sheets) {
    const source = parsed.sheets.find((candidate) => candidate.id === current.id)
      ?? parsed.sheets.find((candidate) => candidate.name.toLocaleLowerCase() === current.name.toLocaleLowerCase());
    if (!source) continue;
    const fontSizesPx: XlsxSheetLayoutRepair['fontSizesPx'] = [];
    for (const [rowKey, sourceRow] of Object.entries(source.cells)) for (const [columnKey, sourceCell] of Object.entries(sourceRow)) {
      const fontSizePx = sourceCell.style?.fontSizePx;
      if (fontSizePx === undefined) continue;
      const row = Number(rowKey);
      const column = Number(columnKey);
      if (current.cells[rowKey]?.[columnKey]?.style?.fontSizePx !== fontSizePx) fontSizesPx.push({ row, column, fontSizePx });
    }
    const defaultChanged = current.defaultRowHeightPx !== source.defaultRowHeightPx || current.defaultColumnWidthPx !== source.defaultColumnWidthPx;
    const rowChanges = changedMapEntries(current.rowHeightsPx ?? {}, source.rowHeightsPx ?? {});
    const columnChanges = changedMapEntries(current.columnWidthsPx ?? {}, source.columnWidthsPx ?? {});
    const paneChanged = JSON.stringify(current.pane) !== JSON.stringify(source.pane);
    if (!defaultChanged && !rowChanges && !columnChanges && !fontSizesPx.length && !paneChanged) continue;
    if (defaultChanged) defaults += 1;
    rows += rowChanges;
    columns += columnChanges;
    fonts += fontSizesPx.length;
    if (paneChanged) panes += 1;
    sheets.push({
      sheetId: current.id,
      sourceSheetName: source.name,
      defaultRowHeightPx: source.defaultRowHeightPx,
      defaultColumnWidthPx: source.defaultColumnWidthPx,
      rowHeightsPx: { ...(source.rowHeightsPx ?? {}) },
      columnWidthsPx: { ...(source.columnWidthsPx ?? {}) },
      pane: structuredClone(source.pane),
      fontSizesPx,
    });
  }
  return { schema: 'XlsxLayoutRepairPlan', sourceChecksum: artifact.checksum, sheets, summary: { sheets: sheets.length, defaults, rows, columns, fonts, panes }, report: structuredClone(imported.report) };
}

function changedMapEntries(current: Record<number, number>, source: Record<number, number>): number {
  const keys = new Set([...Object.keys(current), ...Object.keys(source)]);
  let changed = 0;
  for (const key of keys) if (current[Number(key)] !== source[Number(key)]) changed += 1;
  return changed;
}
