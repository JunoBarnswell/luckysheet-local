#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf run-code "(page) => page.evaluate(async () => ({grid: (() => { const e = document.querySelector('[data-testid=sheet-canvas]'); return e ? {rowCount:e.getAttribute('aria-rowcount'), colCount:e.getAttribute('aria-colcount'), rowIndex:e.getAttribute('aria-rowindex'), colIndex:e.getAttribute('aria-colindex'), active: document.activeElement?.outerHTML?.slice(0,200)} : null; })(), clipboard: await navigator.clipboard.readText(), perm: await navigator.permissions.query({name:'clipboard-read'}).then(p=>p.state)}))"
