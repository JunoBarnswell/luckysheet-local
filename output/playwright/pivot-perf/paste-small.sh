#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
PW=/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh
"$PW" --session pivot-perf eval "navigator.clipboard.writeText(Array.from({length:101},(_,r)=>Array.from({length:20},(_,c)=>r===0 ? (c===0?'Category':c===1?'Value':'Text'+c) : (c===0?'M'+(r%100):c===1?String(r):'T'+(r%10))).join('\\t')).join('\\n')).then(()=> 'clipboard-ready')"
