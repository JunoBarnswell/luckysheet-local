import type { FormulaEngine, ResolvedSpill } from '@react-sheets/formula-engine';

/**
 * A single bounded page from the canonical Rust formula projection.
 * Formula results and spill ranges are owned by the revision-pinned kernel;
 * this module never mutates WorksheetModel or configures a local evaluator.
 */
export interface FormulaSpillPage {
  readonly revision: number;
  readonly spills: readonly ResolvedSpill[];
  readonly nextCursor?: string | null;
}

export interface FormulaSpillPageOptions {
  readonly cursor?: string;
  readonly limit?: number;
}

/** Read one bounded spill page without materializing a second formula model. */
export function readFormulaSpillPage(
  engine: FormulaEngine,
  sheetId: string,
  options: FormulaSpillPageOptions = {},
): FormulaSpillPage {
  return engine.getSpillsForSheetPage(sheetId, options);
}

/** Iterate the canonical spill projection page by page. */
export function* readFormulaSpillPages(
  engine: FormulaEngine,
  sheetId: string,
  options: FormulaSpillPageOptions = {},
): Generator<FormulaSpillPage, void, undefined> {
  let cursor = options.cursor;
  let revision: number | undefined;
  const seenCursors = new Set<string>();
  do {
    if (cursor !== undefined) {
      if (seenCursors.has(cursor)) throw new Error('FORMULA_SPILL_PAGE_CURSOR_REPEATED: canonical spill pagination did not advance');
      seenCursors.add(cursor);
    }
    const page = readFormulaSpillPage(engine, sheetId, { ...options, ...(cursor ? { cursor } : {}) });
    if (revision === undefined) revision = page.revision;
    else if (page.revision !== revision) throw new Error('FORMULA_SPILL_PAGE_REVISION_CHANGED: spill pages crossed workbook revisions');
    yield page;
    cursor = page.nextCursor ?? undefined;
  } while (cursor !== undefined);
}
