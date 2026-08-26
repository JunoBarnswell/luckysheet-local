import fs from 'node:fs';
import path from 'node:path';
import { allAcceptanceCases, commandIdsForAcceptance, validateAcceptanceMatrix } from '../e2e/acceptance-matrix';

const repoRoot = path.resolve(import.meta.dirname, '..');
const requiredSuites = [
  'home-ribbon.spec.ts',
  'insert-ribbon.spec.ts',
  'elastic-sheet-extent.spec.ts',
  'ribbon-visual-parity.spec.ts',
  'permission-matrix.spec.ts',
  'persistence-roundtrip.spec.ts',
].map((file) => path.join(repoRoot, 'e2e', file));

const errors = [...validateAcceptanceMatrix()];
const ribbonSource = fs.readFileSync(path.join(repoRoot, 'apps', 'web', 'src', 'components', 'Ribbon.tsx'), 'utf8');
if (!ribbonSource.includes('data-ribbon-command={id}')) errors.push('CatalogButton does not expose the canonical ribbon command identity');
if (!ribbonSource.includes('data-ribbon-surface={ribbonSurfaceId}')) errors.push('CatalogButton does not expose the canonical ribbon surface identity');
for (const file of requiredSuites) {
  if (!fs.existsSync(file)) errors.push(`missing acceptance suite ${path.relative(repoRoot, file)}`);
}
if (commandIdsForAcceptance().length === 0) errors.push('acceptance matrix has no executable ribbon commands');
if (allAcceptanceCases().some((entry) => entry.id.trim() === '')) errors.push('acceptance matrix contains an empty case id');

if (errors.length > 0) {
  console.error(errors.join('\n'));
  process.exitCode = 1;
} else {
  console.log(`Acceptance matrix check passed: ${allAcceptanceCases().length} cases, ${commandIdsForAcceptance().length} ribbon commands, ${requiredSuites.length} suites.`);
}
