import type { DateSystem } from './types';

/** Excel serial date — 与 JS Date 分离 */
export interface ExcelSerialDate {
  serial: number;
  system: DateSystem;
}

const MS_PER_DAY = 86_400_000;
const EXCEL_EPOCH_1900 = Date.UTC(1899, 11, 30);
const EXCEL_EPOCH_1904 = Date.UTC(1904, 0, 1);

/** 1900 系统闰年 bug: Excel 错误地认为 1900 是闰年 */
export function serialToExcelDate(serial: number, system: DateSystem = '1900'): Date {
  if (system === '1904') {
    return new Date(EXCEL_EPOCH_1904 + serial * MS_PER_DAY);
  }
  let adjusted = serial;
  if (serial >= 60) adjusted += 1;
  return new Date(EXCEL_EPOCH_1900 + adjusted * MS_PER_DAY);
}

export function dateToSerial(date: Date, system: DateSystem = '1900'): number {
  const epoch = system === '1904' ? EXCEL_EPOCH_1904 : EXCEL_EPOCH_1900;
  let serial = (date.getTime() - epoch) / MS_PER_DAY;
  if (system === '1900' && serial >= 60) serial -= 1;
  return serial;
}

export function parseDateSystem(workbookXml?: string): DateSystem {
  if (workbookXml?.includes('<workbookPr date1904="1"')) return '1904';
  return '1900';
}
