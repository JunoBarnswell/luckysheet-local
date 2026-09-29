export function parseCellMatrixCoordinate(key: string, axis: 'row' | 'column', source: 'deferred' | 'JSON import' | 'workbook snapshot'): number {
  const coordinate = Number(key);
  if (!Number.isSafeInteger(coordinate) || coordinate < 0 || String(coordinate) !== key) {
    throw new Error(`CELL_MATRIX_INVALID_COORDINATE: ${source} ${axis} key ${JSON.stringify(key)} is not a canonical non-negative integer`);
  }
  return coordinate;
}
