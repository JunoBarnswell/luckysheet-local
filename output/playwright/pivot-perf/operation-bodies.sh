#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
CLI=/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh
for i in 69 72 74 76; do echo REQUEST_$i; $CLI --session pivot-perf request-body $i | head -c 600; echo; done
