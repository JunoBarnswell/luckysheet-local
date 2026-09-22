#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf run-code "async (page) => { await page.getByRole('textbox', { name: '用户名' }).fill('perfadmin'); await page.getByRole('textbox', { name: '密码' }).fill('PerfPass-2026-Strong!'); await page.getByRole('button', { name: '登录' }).click(); await page.waitForURL(/\\/workbooks/); return page.url(); }"
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf snapshot
