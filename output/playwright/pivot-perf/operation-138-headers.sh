#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
CLI=/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh
$CLI --session pivot-perf response-headers 138
$CLI --session pivot-perf request-headers 138
