#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
CLI=/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh
$CLI --session pivot-perf click e379
$CLI --session pivot-perf wait 1000
$CLI --session pivot-perf snapshot
