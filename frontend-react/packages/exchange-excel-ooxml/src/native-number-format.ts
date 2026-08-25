import type { WorkbookSnapshot } from '@react-sheets/core-model';

/** Excel's built-in number-format codes used by cell styles and Pivot dataFields. */
const BUILT_IN_NUMBER_FORMATS: ReadonlyMap<number, string> = new Map([
  [0, 'General'], [1, '0'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00'],
  [9, '0%'], [10, '0.00%'], [11, '0.00E+00'], [12, '# ?/?'], [13, '# ??/??'],
  [14, 'm/d/yy'], [15, 'd-mmm-yy'], [16, 'd-mmm'], [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'], [19, 'h:mm:ss AM/PM'], [20, 'h:mm'], [21, 'h:mm:ss'], [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'], [38, '#,##0 ;[Red](#,##0)'], [39, '#,##0.00;(#,##0.00)'], [40, '#,##0.00;[Red](#,##0.00)'],
  [41, '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)'], [42, '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)'],
  [43, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'], [44, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'],
  [45, 'mm:ss'], [46, '[h]:mm:ss'], [47, 'mmss.0'], [48, '##0.0E+0'], [49, '@'],
]);

export function builtInNumberFormat(id: number): string | undefined {
  return BUILT_IN_NUMBER_FORMATS.get(id);
}

export function builtInNumberFormatId(format: string): number | undefined {
  for (const [id, code] of BUILT_IN_NUMBER_FORMATS) if (code === format) return id;
  return undefined;
}

/**
 * Allocate custom IDs in the same deterministic order as styles.xml: cells,
 * named cell styles, then canonical Pivot value fields.
 */
export function collectCustomNumberFormatIds(snapshot: WorkbookSnapshot): Map<string, number> {
  const custom = new Map<string, number>();
  let next = 164;
  const add = (format: string | undefined) => {
    if (!format || builtInNumberFormatId(format) !== undefined || custom.has(format)) return;
    custom.set(format, next);
    next += 1;
  };
  for (const sheet of snapshot.sheets) {
    for (const row of Object.values(sheet.cells)) for (const cell of Object.values(row)) add(cell.numberFormat);
  }
  for (const template of snapshot.cellStyleTemplates ?? []) add(template.style.numberFormat);
  for (const sheet of snapshot.sheets) for (const pivot of sheet.pivots) for (const value of pivot.layout.values) add(value.numberFormat);
  return custom;
}

export function numberFormatId(format: string, custom: ReadonlyMap<string, number>): number | undefined {
  return builtInNumberFormatId(format) ?? custom.get(format);
}
