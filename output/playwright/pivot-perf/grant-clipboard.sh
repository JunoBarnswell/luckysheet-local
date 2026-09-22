#!/usr/bin/env bash
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf run-code "(page) => page.context().grantPermissions(['clipboard-read','clipboard-write'], {origin: 'http://127.0.0.1:8083'}).then(() => 'clipboard-permissions-granted')"
