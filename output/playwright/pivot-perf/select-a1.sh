#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf run-code "async (page) => { const box = page.getByRole('textbox', { name: '选中单元格' }); await box.fill('A1'); await box.press('Enter'); return await box.inputValue(); }"
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf snapshot
