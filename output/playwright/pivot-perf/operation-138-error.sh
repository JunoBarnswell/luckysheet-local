#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
CLI=/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh
$CLI --session pivot-perf response-body 138 | head -c 3000
