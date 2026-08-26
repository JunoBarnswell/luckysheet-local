import fs from 'node:fs';
import path from 'node:path';
import { allAcceptanceCases, commandIdsForAcceptance, CORE_INTERACTION_CASES, RIBBON_LAYOUT_CASES, RIBBON_VISUAL_GOLDEN_CASES, validateAcceptanceMatrix } from '../e2e/acceptance-matrix';

const repoRoot = path.resolve(import.meta.dirname, '..');
const requiredSuites = [
  'home-ribbon.spec.ts',
  'insert-ribbon.spec.ts',
  'ribbon-layout-matrix.spec.ts',
  'elastic-sheet-extent.spec.ts',
  'designer-visual.spec.ts',
  'ribbon-visual-parity.spec.ts',
  'permission-matrix.spec.ts',
  'persistence-roundtrip.spec.ts',
  'workbook-hub.spec.ts',
].map((file) => path.join(repoRoot, 'e2e', file));

const errors = [...validateAcceptanceMatrix()];
const ribbonSource = fs.readFileSync(path.join(repoRoot, 'apps', 'web', 'src', 'components', 'Ribbon.tsx'), 'utf8');
if (!ribbonSource.includes('data-ribbon-command={id}')) errors.push('CatalogButton does not expose the canonical ribbon command identity');
if (!ribbonSource.includes('data-ribbon-surface={ribbonSurfaceId}')) errors.push('CatalogButton does not expose the canonical ribbon surface identity');
if (!ribbonSource.includes('data-ribbon-layout-node={ribbonLayoutNodeId}')) errors.push('CatalogButton does not expose the canonical ribbon layout-node identity');
for (const file of requiredSuites) {
  if (!fs.existsSync(file)) errors.push(`missing acceptance suite ${path.relative(repoRoot, file)}`);
}
if (commandIdsForAcceptance().length === 0) errors.push('acceptance matrix has no executable ribbon commands');
if (allAcceptanceCases().some((entry) => entry.id.trim() === '')) errors.push('acceptance matrix contains an empty case id');
for (const entry of CORE_INTERACTION_CASES) {
  if (!fs.existsSync(path.join(repoRoot, entry.sourceSpec))) errors.push(`missing core interaction source ${entry.sourceSpec}`);
}
if (new Set(RIBBON_VISUAL_GOLDEN_CASES.map((entry) => entry.screenshot)).size !== RIBBON_VISUAL_GOLDEN_CASES.length) errors.push('visual golden screenshot paths must be unique');
if (new Set(RIBBON_VISUAL_GOLDEN_CASES.map((entry) => entry.shellScreenshot)).size !== RIBBON_VISUAL_GOLDEN_CASES.length / 5) errors.push('shell visual golden paths must be unique per locale and viewport');

if (errors.length > 0) {
  console.error(errors.join('\n'));
  process.exitCode = 1;
} else {
  console.log(`Acceptance matrix check passed: ${allAcceptanceCases().length} cases, ${commandIdsForAcceptance().length} ribbon commands, ${CORE_INTERACTION_CASES.length} core interactions, ${RIBBON_LAYOUT_CASES.length} layout targets, ${RIBBON_VISUAL_GOLDEN_CASES.length} visual goldens, ${requiredSuites.length} suites.`);
}
