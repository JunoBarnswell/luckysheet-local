#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf run-code "async (page) => { await page.waitForFunction(() => Array.from(document.querySelectorAll('[role=status]')).some((element) => element.textContent?.includes('已保存')), { timeout: 180000 }); return await page.getByRole('status').allTextContents(); }"
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf snapshot
