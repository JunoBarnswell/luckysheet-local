export interface KeyTipState {
  active: boolean;
  prefix: string;
}

export interface KeyTipBinding {
  sequence: string;
  target: { kind: 'tab' | 'command'; id: string };
}

export interface KeyTipTransition {
  state: KeyTipState;
  action?: KeyTipBinding['target'];
}

export const INITIAL_KEY_TIP_STATE: KeyTipState = { active: false, prefix: '' };

/**
 * Excel's first-level access letters and the stable command access letters
 * used by the workbook Ribbon. Sequences are declarative so the shell and
 * keyboard resolver never invent different navigation trees.
 */
export const EXCEL_KEY_TIP_BINDINGS: readonly KeyTipBinding[] = [
  { sequence: 'H', target: { kind: 'tab', id: 'home' } },
  { sequence: 'N', target: { kind: 'tab', id: 'insert' } },
  { sequence: 'P', target: { kind: 'tab', id: 'pageLayout' } },
  { sequence: 'M', target: { kind: 'tab', id: 'formulas' } },
  { sequence: 'A', target: { kind: 'tab', id: 'data' } },
  { sequence: 'R', target: { kind: 'tab', id: 'review' } },
  { sequence: 'W', target: { kind: 'tab', id: 'view' } },
  { sequence: 'F', target: { kind: 'tab', id: 'file' } },
  { sequence: 'HV', target: { kind: 'command', id: 'paste' } },
  { sequence: 'HX', target: { kind: 'command', id: 'cut' } },
  { sequence: 'HCP', target: { kind: 'command', id: 'copy' } },
  { sequence: 'HFP', target: { kind: 'command', id: 'format-painter' } },
  { sequence: 'HB', target: { kind: 'command', id: 'bold' } },
  { sequence: 'HI', target: { kind: 'command', id: 'italic' } },
  { sequence: 'HU', target: { kind: 'command', id: 'underline' } },
  { sequence: 'HFC', target: { kind: 'command', id: 'font-color' } },
  { sequence: 'HFF', target: { kind: 'command', id: 'fill-color' } },
  { sequence: 'HOO', target: { kind: 'command', id: 'orientation-menu' } },
  { sequence: 'HWM', target: { kind: 'command', id: 'wrapText' } },
  { sequence: 'HWS', target: { kind: 'command', id: 'shrinkToFit' } },
  { sequence: 'HMC', target: { kind: 'command', id: 'mergeCenter' } },
  { sequence: 'HFD', target: { kind: 'command', id: 'fillDown' } },
  { sequence: 'HFR', target: { kind: 'command', id: 'fillRight' } },
  { sequence: 'HFE', target: { kind: 'command', id: 'fillSeries' } },
  { sequence: 'HFL', target: { kind: 'command', id: 'fillLeft' } },
  { sequence: 'HCF', target: { kind: 'command', id: 'conditionalFormat' } },
  { sequence: 'HFS', target: { kind: 'command', id: 'sortRange' } },
  { sequence: 'HFG', target: { kind: 'command', id: 'findReplace' } },
  { sequence: 'HGO', target: { kind: 'command', id: 'goTo' } },
  { sequence: 'HCC', target: { kind: 'command', id: 'clearContents' } },
  { sequence: 'HCL', target: { kind: 'command', id: 'clearFormats' } },
  { sequence: 'HCA', target: { kind: 'command', id: 'clearAll' } },
  { sequence: 'HCR', target: { kind: 'command', id: 'clearCommentsNotes' } },
  { sequence: 'HCH', target: { kind: 'command', id: 'clearHyperlinks' } },
  { sequence: 'NT', target: { kind: 'command', id: 'worksheetTable' } },
  { sequence: 'NP', target: { kind: 'command', id: 'picture' } },
  { sequence: 'NSH', target: { kind: 'command', id: 'shapesLines' } },
  { sequence: 'NSS', target: { kind: 'command', id: 'screenshot' } },
  { sequence: 'NC', target: { kind: 'command', id: 'chartBuilder' } },
  { sequence: 'NR', target: { kind: 'command', id: 'recommendedCharts' } },
  { sequence: 'NK', target: { kind: 'command', id: 'sparkline' } },
  { sequence: 'NI', target: { kind: 'command', id: 'icons' } },
  { sequence: 'NM', target: { kind: 'command', id: 'models3d' } },
  { sequence: 'NA', target: { kind: 'command', id: 'smartArt' } },
  { sequence: 'NSC', target: { kind: 'command', id: 'camera' } },
  { sequence: 'NL', target: { kind: 'command', id: 'hyperlink' } },
  { sequence: 'NCO', target: { kind: 'command', id: 'threadedComment' } },
  { sequence: 'NE', target: { kind: 'command', id: 'equation' } },
  { sequence: 'NYS', target: { kind: 'command', id: 'symbol' } },
];

function normalizeKey(key: string): string {
  return key.length === 1 ? key.toLocaleUpperCase() : key;
}

export function keyTipTransition(state: KeyTipState, key: string, bindings: readonly KeyTipBinding[] = EXCEL_KEY_TIP_BINDINGS): KeyTipTransition {
  if (!state.active) return { state };
  if (key === 'Escape') return { state: INITIAL_KEY_TIP_STATE };
  const nextPrefix = `${state.prefix}${normalizeKey(key)}`;
  const candidates = bindings.filter((binding) => binding.sequence.startsWith(nextPrefix));
  if (candidates.length === 0) return { state: INITIAL_KEY_TIP_STATE };
  const exact = candidates.find((binding) => binding.sequence === nextPrefix);
  const hasChildren = candidates.some((binding) => binding.sequence.length > nextPrefix.length);
  if (exact && (exact.target.kind === 'tab' || !hasChildren)) {
    return { state: exact.target.kind === 'tab' ? { active: true, prefix: nextPrefix } : INITIAL_KEY_TIP_STATE, action: exact.target };
  }
  return { state: { active: true, prefix: nextPrefix } };
}

export function keyTipCandidates(prefix: string, bindings: readonly KeyTipBinding[] = EXCEL_KEY_TIP_BINDINGS): readonly KeyTipBinding[] {
  const normalized = prefix.toLocaleUpperCase();
  return bindings.filter((binding) => binding.sequence.startsWith(normalized));
}
