const INVALID_WORKSHEET_NAME_CHARACTERS = new Set(['/', '\\', '?', '*', ':', '[', ']']);

export function assertCanonicalWorksheetName(name: unknown): asserts name is string {
  if (typeof name !== 'string' || !name.trim() || name.length > 31
    || name.startsWith("'") || name.endsWith("'") || name.toUpperCase() === 'HISTORY'
    || [...name].some((character) => INVALID_WORKSHEET_NAME_CHARACTERS.has(character))) {
    throw new Error('Worksheet name is invalid under Excel naming rules');
  }
}
