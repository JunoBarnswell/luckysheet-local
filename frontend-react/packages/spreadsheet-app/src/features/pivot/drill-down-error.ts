export class PivotDrillDownError extends Error {
  readonly code = 'UNSUPPORTED_FEATURE';

  constructor() {
    super('Pivot drill-down cannot combine worksheet-range provenance with an overlapping block-backed region; use the canonical data-source Pivot source');
    this.name = 'PivotDrillDownError';
  }
}
