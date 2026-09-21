#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
CLI=/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh
$CLI --session pivot-perf request-body 138 > operation-138-body.json
wc -c operation-138-body.json
head -c 180 operation-138-body.json; echo
