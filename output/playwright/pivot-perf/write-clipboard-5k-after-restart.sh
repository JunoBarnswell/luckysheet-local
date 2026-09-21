#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf run-code "(page) => page.evaluate(async () => { const rows = 5000; const columns = 20; const lines = []; const header = []; for (let c = 0; c < columns; c++) header.push(c === 0 ? 'Category' : c === 1 ? 'Value' : 'Text' + String(c)); lines.push(header.join('\\t')); for (let r = 1; r < rows; r++) { const row = []; for (let c = 0; c < columns; c++) row.push(c === 0 ? 'M' + String(r % 100) : c === 1 ? String(r) : 'T' + String(r % 32)); lines.push(row.join('\\t')); } const text = lines.join('\\r\\n'); return navigator.clipboard.writeText(text).then(() => ({ rows, columns, bytes: text.length })); })"
