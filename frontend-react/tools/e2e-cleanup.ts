import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import type { FullConfig } from '@playwright/test';

const frontendRoot = path.resolve(import.meta.dirname, '..');
const connectedWorkbookLedgerPath = path.join(frontendRoot, 'test-results', 'connected-workbooks.ndjson');

interface OidcUserPayload {
  readonly access_token?: unknown;
}

function accessToken(): string | null {
  const userFile = process.env.E2E_OIDC_USER_FILE?.trim();
  if (!userFile) return null;
  const payload = JSON.parse(fs.readFileSync(path.resolve(process.cwd(), userFile), 'utf8')) as OidcUserPayload;
  if (typeof payload.access_token !== 'string' || payload.access_token.trim() === '') {
    throw new Error('E2E cleanup requires access_token in E2E_OIDC_USER_FILE');
  }
  return payload.access_token;
}

function recordedUnitIds(): readonly string[] {
  if (!fs.existsSync(connectedWorkbookLedgerPath)) return [];
  const unitIds = fs.readFileSync(connectedWorkbookLedgerPath, 'utf8')
    .split(/\r?\n/)
    .filter(Boolean)
    .map((line) => JSON.parse(line) as { unitId?: unknown })
    .map((entry) => typeof entry.unitId === 'string' ? entry.unitId : '')
    .filter(Boolean);
  return [...new Set(unitIds)];
}

async function requireNoContent(url: string, token: string): Promise<void> {
  const response = await fetch(url, { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } });
  if (response.status !== 204) {
    throw new Error(`E2E workbook cleanup failed: DELETE ${url} returned ${response.status} ${await response.text()}`);
  }
}

export default async function globalTeardown(_config: FullConfig): Promise<void> {
  const unitIds = recordedUnitIds();
  if (unitIds.length === 0) return;
  const token = accessToken();
  if (!token) throw new Error('E2E created connected workbooks but E2E_OIDC_USER_FILE is unavailable for cleanup');
  const apiOrigin = (process.env.REACT_SHEETS_API_ORIGIN ?? 'http://127.0.0.1:8082').replace(/\/$/, '');
  const failures: string[] = [];
  for (const unitId of unitIds) {
    const workbookUrl = `${apiOrigin}/api/workbooks/${encodeURIComponent(unitId)}`;
    try {
      await requireNoContent(workbookUrl, token);
      await requireNoContent(`${workbookUrl}/purge`, token);
    } catch (error) {
      failures.push(error instanceof Error ? error.message : String(error));
    }
  }
  if (failures.length > 0) throw new Error(`Connected workbook cleanup failed:\n${failures.join('\n')}`);
}
