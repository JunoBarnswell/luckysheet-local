import {
  createUniver,
  defaultTheme,
  LocaleType,
  LogLevel,
  mergeLocales,
} from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US';
import { UniverSheetsAdvancedPreset } from '@univerjs/preset-sheets-advanced';
import sheetsAdvancedEnUS from '@univerjs/preset-sheets-advanced/locales/en-US';
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing';
import sheetsDrawingEnUS from '@univerjs/preset-sheets-drawing/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-advanced/lib/index.css';
import '@univerjs/preset-sheets-drawing/lib/index.css';

const { univerAPI } = createUniver({
  locale: LocaleType.EN_US,
  locales: {
    [LocaleType.EN_US]: mergeLocales(sheetsCoreEnUS, sheetsAdvancedEnUS, sheetsDrawingEnUS),
  },
  logLevel: LogLevel.WARN,
  theme: defaultTheme,
  presets: [
    UniverSheetsCorePreset({
      container: 'univer',
      header: true,
      toolbar: true,
      formulaBar: true,
      footer: true,
      workerURL: new Worker(new URL('./worker.js', import.meta.url), {
        type: 'module',
      }),
    }),
    UniverSheetsDrawingPreset(),
    UniverSheetsAdvancedPreset({
      useWorker: true,
      universerEndpoint: window.location.origin,
      exchangeClientOptions: {
        minSheetRowCount: 50,
        minSheetColumnCount: 12,
      },
    }),
  ],
});

const workbook = univerAPI.createWorkbook({ id: 'pro-ui-research', name: 'Pro UI Research' });
const sheet = workbook.getActiveSheet();

sheet.getRange('A1:D6').setValues([
  ['Univer Pro UI research', 'Value', 'Formula', 'Status'],
  ['Chart source', 120, '=B2*1.2', 'Chart/Pivot/Print preset loaded'],
  ['Second row', 240, '=B3*1.2', 'Formula engine active'],
  ['Third row', 360, '=B4*1.2', 'Canvas sheet active'],
  ['Fourth row', 480, '=B5*1.2', 'React UI surface active'],
  ['Fifth row', 600, '=B6*1.2', 'License gate intentionally unconfigured'],
]);

sheet.getRange('A1:D1').setFontWeight('bold');
sheet.getRange('A1:D6').setHorizontalAlignment('center');
sheet.getRange('B2:C6').setNumberFormat('0.00');
sheet.getRange('A1:D6').setWrap(true);

window.__univerProResearch = { univerAPI, workbook, sheet };
