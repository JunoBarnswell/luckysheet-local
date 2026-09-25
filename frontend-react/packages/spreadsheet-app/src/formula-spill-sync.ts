import type { WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
import type { FormulaEngine, SheetTableRef, SpillEnvironment } from '@react-sheets/formula-engine';

export function createSpillEnvironment(sheet: WorksheetModel): SpillEnvironment {
  const blockedRanges = [
    ...sheet.merges.map(({ range }) => ({
      startRow: range.startRow,
      endRow: range.endRow,
      startColumn: range.startColumn,
      endColumn: range.endColumn,
    })),
    ...sheet.sheetTables.map(({ range }) => ({
      startRow: range.startRow,
      endRow: range.endRow,
      startColumn: range.startColumn,
      endColumn: range.endColumn,
    })),
  ];
  const environment: SpillEnvironment = {
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    isOccupied: (row, column) => {
      const cell = sheet.cells.get(row, column);
      if (cell && (cell.formula !== undefined || (cell.value != null && cell.value !== ''))) return true;
      return false;
    },
    getBlockedRanges: () => blockedRanges,
    ensureExtent: (rowCount, columnCount) => {
      sheet.ensureRangeExtent(0, Math.max(0, rowCount - 1), 0, Math.max(0, columnCount - 1));
      environment.rowCount = sheet.rowCount;
      environment.columnCount = sheet.columnCount;
    },
  };
  return environment;
}

export function configureFormulaSpillEnvironment(engine: FormulaEngine, sheet: WorksheetModel): void {
  engine.setSpillEnvironment(sheet.id, createSpillEnvironment(sheet));
}

export function syncFormulaSpillsToSheet(engine: FormulaEngine, sheet: WorksheetModel): void {
  const spills = engine.getSpillsForSheet(sheet.id);
  sheet.spillRanges.splice(0, sheet.spillRanges.length, ...spills);
}

export function configureWorkbookSpillEnvironments(engine: FormulaEngine, workbook: WorkbookModel): void {
  for (const sheet of workbook.getSheets()) {
    configureFormulaSpillEnvironment(engine, sheet);
  }
}

export function workbookSheetTables(workbook: WorkbookModel): SheetTableRef[] {
  const tables: SheetTableRef[] = [];
  for (const sheet of workbook.getSheets()) {
    for (const table of sheet.sheetTables) {
      tables.push({
        id: table.id,
        sheetId: table.sheetId,
        name: table.name,
        range: table.range,
        hasHeaderRow: table.hasHeaderRow,
        hasTotalRow: table.hasTotalRow,
        columns: table.columns.map((column) => ({ id: column.id, name: column.name })),
      });
    }
  }
  return tables;
}

export function syncWorkbookSheetTables(engine: FormulaEngine, workbook: WorkbookModel, recalculate = true): void {
  engine.setSheetTables(workbookSheetTables(workbook), recalculate);
}
