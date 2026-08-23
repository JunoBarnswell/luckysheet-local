export function columnLabel(index: number): string {
  let value = index + 1;
  let label = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

export function parseColumnLabel(label: string): number {
  let column = 0;
  for (const character of label.toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return column - 1;
}

export function cellAddress(row: number, column: number): string {
  return `${columnLabel(column)}${row + 1}`;
}

export function parseAddress(address: string): { column: number; row: number } | undefined {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
  if (!match?.[1] || !match[2]) return undefined;
  return { column: parseColumnLabel(match[1]), row: Number(match[2]) - 1 };
}

export function a1Range(startRow: number, startColumn: number, endRow: number, endColumn: number): string {
  return `${cellAddress(startRow, startColumn)}:${cellAddress(endRow, endColumn)}`;
}
