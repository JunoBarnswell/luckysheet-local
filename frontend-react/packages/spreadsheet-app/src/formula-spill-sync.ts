import type { WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
import type { FormulaEngine, SheetTableRef, SpillEnvironment } from '@react-sheets/formula-engine';

export function createSpillEnvironment(sheet: WorksheetModel): SpillEnvironment {
  return {
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    isOccupied: (row, column) => {
      const cell = sheet.cells.get(row, column);
      return Boolean(cell && (cell.formula !== undefined || (cell.value != null && cell.value !== '')));
    },
  };
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

export function syncWorkbookSpills(engine: FormulaEngine, workbook: WorkbookModel): void {
  for (const sheet of workbook.getSheets()) {
    syncFormulaSpillsToSheet(engine, sheet);
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

export function syncWorkbookSheetTables(engine: FormulaEngine, workbook: WorkbookModel): void {
  engine.setSheetTables(workbookSheetTables(workbook));
}
