# Issue #304 Design QA

final result: passed

## Compact-width alignment follow-up

- `ribbonLayoutModeForWidth(1183)` now reaches the HOME compact geometry rather than reusing the 1783px wide geometry.
- Compact HOME group widths are Clipboard 124px, Font 224px, Alignment 248px, Number 154px, Styles 220px, Cells 144px, Editing 300px.
- Compact Font controls use a 100px family field, 45px size field and 25px A↑/A↓ controls; Alignment's wrap/merge controls are 74px/84px with 22px alignment glyph buttons.
- Important padding utilities are explicit (`!px-*`) so the shared Button base padding cannot compress icon boxes; measured compact Fluent glyphs are 16px (font) and 14px (alignment).
- Browser comment target at 1183px was rechecked with no group-child overflow or overlap.

## Icon alignment follow-up

- Official Fluent regular assets are vendored under `apps/web/public/icons/fluent/` and used by HOME/INSERT command-bar surfaces.
- HOME: Paste 64×104 with 32 px glyph; font family 214 px; number format 188 px; compact command glyphs explicit 16 px.
- INSERT: Illustration items are six fixed 50 px cells inside the 308 px group, preventing Screenshot/SmartArt overlap. Charts use a fixed 320 px grid, 48 px family cells, 20 px thumbnails and a 284 px categorized gallery. Sparklines use three equal 58 px cells inside the 181 px group.
- Latest browser evidence (temporary files): `C:\Users\kuo13\AppData\Local\Temp\issue-304-icon-home-1783.png`, `C:\Users\kuo13\AppData\Local\Temp\issue-304-icon-gallery-final.png`, and `C:\Users\kuo13\AppData\Local\Temp\issue-304-icon-alignment-insert.png`.

## Reference

- HOME: `C:\Users\kuo13\AppData\Local\Temp\codex-clipboard-fe420d56-7b06-41bf-801d-250843c45638.png` (1783×141)
- INSERT: `C:\Users\kuo13\AppData\Local\Temp\codex-clipboard-c9bac183-6f0e-4b48-a9e6-c1ca1048dabe.png` (1905×135)

## Implemented visual contract

- HOME order: Clipboard, Font, Alignment, Number, Styles, Cells, Editing.
- INSERT order: Tables, Illustrations, Controls, Charts, Sparklines, Filters, Links, Comments, Text, Symbols.
- Wide thresholds include both reference widths.
- Command content height is 135px; HOME and INSERT wide minimum widths match their reference images.
- Group widths were measured from the supplied separators and encoded in the shared layout renderer.
- HOME no longer renders History or direct Strikethrough surfaces.
- Paste is a real split surface; Format Cells launchers target Font, Alignment, and Number.

## Automated gates

- TypeScript: passed.
- Unit tests: 737/737 passed.
- Architecture boundaries and mutation registry: passed.
- Contracts and acceptance matrix: passed.
- Production build: passed with pre-existing circular-chunk and large-chunk warnings.
- `git diff --check`: passed.

## In-app browser evidence

- Page identity: `http://127.0.0.1:4180/workbooks/wb-d6c35ee4-9a3e-4283-95d5-ec0b370ae12a?initialCell=B1`, title `React Sheets`.
- HOME browser capture at 1783×900 viewport: [issue-304-home-final.png](C:/Users/kuo13/AppData/Local/Temp/issue-304-home-final.png); Ribbon content rect is 1783.2×135 px.
- INSERT browser capture at 1905×900 viewport: [issue-304-insert-final.png](C:/Users/kuo13/AppData/Local/Temp/issue-304-insert-final.png); Ribbon content rect is 1905.6×135 px.
- HOME DOM exposes the seven required groups and the paste split, format painter, phonetic guide, and three Format Cells launchers. INSERT DOM exposes the ten required groups, all chart families/submenus, symbol, comments, text, controls, and disabled Pivot-context controls when no Pivot is selected.
- Font launcher opened `设置单元格格式` with `字体` selected. Recommended Charts correctly rejected a blank selection with `INVALID_CHART_SOURCE` instead of fabricating a candidate. Symbol opened the Unicode picker. Console readback after reload contained no warning/error entries.

## Component-level Pixel Difference Report

- Structure: PASS — group order, separators, labels, and wide breakpoints match the supplied HOME/INSERT contracts.
- Spacing/layout: PASS — shared renderer uses the measured 1783/1905 target widths and 135 px content height.
- Typography: PASS — localized Chinese labels are rendered by the existing Noto Sans SC/Segoe UI stack without fallback glyph errors.
- Icon semantics: PASS — surfaces use the documented Fluent literal-metaphor mapping; no emoji or placeholder glyphs are introduced.
- Interaction states: PASS — split menus, dialog launchers, disabled Pivot-context controls, and fail-close recommendation behavior were exercised.
- P0/P1 mismatches: none observed in the final contract capture.
- P2 note: the supplied Excel screenshots use Microsoft Office's proprietary multicolor glyph artwork; this repository uses the documented Fluent monoline semantic assets/paths and shared currentColor tokens, so exact proprietary pixel identity is outside the local asset boundary.
