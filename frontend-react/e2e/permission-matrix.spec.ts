import { expect, test } from '@playwright/test';
import { PERMISSION_ENTRYPOINTS, PERMISSION_MATRIX_CASES, PERMISSION_ROLES, PROTECTION_STATES } from './acceptance-matrix';

test.describe('Permission acceptance matrix contract', () => {
  test('declares every role, protection state, and entry point exactly once', () => {
    const expected = PERMISSION_ROLES.length * PROTECTION_STATES.length * PERMISSION_ENTRYPOINTS.length;
    expect(PERMISSION_MATRIX_CASES).toHaveLength(expected);
    expect(new Set(PERMISSION_MATRIX_CASES.map((entry) => entry.id)).size).toBe(expected);
    for (const entry of PERMISSION_MATRIX_CASES) {
      expect(PERMISSION_ROLES).toContain(entry.role);
      expect(PROTECTION_STATES).toContain(entry.protection);
      expect(PERMISSION_ENTRYPOINTS).toContain(entry.entrypoint);
      expect(entry.layers).toEqual(expect.arrayContaining(['contract', 'browser', 'parity']));
    }
  });
});
