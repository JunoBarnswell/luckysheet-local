#!/usr/bin/env bash
set -e
cd /c/Users/kuo13/Projects/luckysheet-local/output/playwright/pivot-perf
/c/Users/kuo13/.codex/skills/playwright/scripts/playwright_cli.sh --session pivot-perf eval "(()=>{const text=Array.from({length:101},(_,r)=>Array.from({length:20},(_,c)=>r===0?(c===0?'Category':c===1?'Value':'Text'+c):(c===0?'M'+(r%100):c===1?String(r):'T'+(r%10))).join('\\t')).join('\\n');const data=new DataTransfer();data.setData('text/plain',text);const event=new ClipboardEvent('paste',{bubbles:true,cancelable:true,clipboardData:data});document.querySelector('[data-testid=sheet-canvas]')?.dispatchEvent(event);return {length:text.length,defaultPrevented:event.defaultPrevented}})()"
