#!/usr/bin/env bash
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf eval "navigator.clipboard.readText().then(t=>({length:t.length,head:t.slice(0,30)}))"
