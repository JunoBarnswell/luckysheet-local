/** Maps one zero-based axis coordinate through an insertion or deletion. */
export function mapAxisCoordinate(
  coordinate: number,
  at: number,
  count: number,
  direction: 1 | -1,
): number | null {
  if (direction === 1) return coordinate >= at ? coordinate + count : coordinate;
  if (coordinate < at) return coordinate;
  if (coordinate < at + count) return null;
  return coordinate - count;
}
