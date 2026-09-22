#!/usr/bin/env bash
set +x
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
TOKEN="$(cat /c/Users/kuo13/AppData/Local/Temp/react-sheets-perf-20260921-181317/bootstrap-token)"
PW=/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh
"$PW" --session pivot-perf fill e11 "$TOKEN" > init.log 2>&1
"$PW" --session pivot-perf fill e14 perfadmin >> init.log 2>&1
"$PW" --session pivot-perf fill e17 'Pivot Performance Admin' >> init.log 2>&1
"$PW" --session pivot-perf fill e20 'PerfPass-2026-Strong!' >> init.log 2>&1
"$PW" --session pivot-perf click e21 >> init.log 2>&1
code=$?
echo "exit=$code"
exit "$code"
