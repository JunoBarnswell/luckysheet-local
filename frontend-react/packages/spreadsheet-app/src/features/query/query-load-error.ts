export class QueryLoadError extends Error {
  constructor(
    readonly code: 'QUERY_LOAD_IN_PROGRESS' | 'QUERY_LOAD_STALE' | 'QUERY_LOAD_CANCELLED',
    readonly queryId: string,
    message: string,
  ) {
    super(`${code}: ${queryId}: ${message}`);
    this.name = 'QueryLoadError';
  }
}
