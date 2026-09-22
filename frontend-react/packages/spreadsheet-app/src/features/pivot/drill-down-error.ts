export class PivotDrillDownError extends Error {
  readonly code = 'UNSUPPORTED_FEATURE';

  constructor() {
    super('Pivot drill-down requires canonical block reads for this data source; no detail sheet was created');
    this.name = 'PivotDrillDownError';
  }
}
