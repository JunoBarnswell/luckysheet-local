# SDK 产品验收：完整单元测试失败清单

2026-10-03，1c9bae6b81ee23f4e512fb9f66c8ca9343d7b04b；`npm run test:unit` 实际执行。1538 tests，1478 Pass，60 Fail。后续改动只涉及浏览器验收键盘输入、CI、验收文档，未改变单元测试实现。完整日志 `/tmp/sdk-all-unit-current.log`。基线为 83 Fail，不能未经逐项对照将当前问题全部归于基线。此清单不意味着错误都在测试；需要逐项判断实际缺陷与过期/非 canonical fixture，不能通过弱化授权或转换约束解决。

| # | 失败测试 | 实际文件/位置 |
|---|---|---|
| 1 | keeps the same wide group widths at every viewport | `apps/web/src/components/home-ribbon-layout.test.ts:10:3` |
| 2 | CommandRuntime keeps formula-rule owners synchronized with a provided FormulaEngine index | `packages/command-runtime/src/index.test.ts:43:1` |
| 3 | CommandRuntime finishes committed replay before fail-stopping on a participant failure | `packages/command-runtime/src/index.test.ts:212:1` |
| 4 | CommandRuntime adopts committed sheet-rename owner facts for undo history | `packages/command-runtime/src/index.test.ts:684:1` |
| 5 | CommandRuntime replays exact structural range-owner facts through undo and redo | `packages/command-runtime/src/index.test.ts:790:1` |
| 6 | CommandRuntime applies workbook-table and data-source range deltas atomically | `packages/command-runtime/src/index.test.ts:987:1` |
| 7 | CellMatrix visits all persisted sparse cells without hydration and propagates reader failures | `packages/core-model/src/index.test.ts:588:1` |
| 8 | CellMatrix applies sparse additions, replacements and deletions without loading an inactive sheet | `packages/core-model/src/index.test.ts:675:1` |
| 9 | CellMatrix keeps deferred cells intact when normalization fails during hydration | `packages/core-model/src/index.test.ts:760:1` |
| 10 | defers sparse worksheet cell hydration across point reads | `packages/core-model/src/index.test.ts:1334:1` |
| 11 | preserves supported pie subtypes and fails closed on unmodeled native splits and rotations | `packages/exchange-excel-ooxml/src/exchange-excel-ooxml.test.ts:16:3` |
| 12 | rejects custom metadata drawing ranges that exceed bounded render work | `packages/exchange-excel-ooxml/src/exchange-excel-ooxml.test.ts:386:3` |
| 13 | reads and rewrites native Pivot cache/table relationship graphs | `packages/exchange-excel-ooxml/src/exchange-excel-ooxml.test.ts:1164:3` |
| 14 | honors explicit OOXML export options before reusing untouched source bytes | `packages/exchange-excel-ooxml/src/exchange-excel-ooxml.test.ts:1941:3` |
| 15 | range move and style preset are atomic and reversible | `packages/sheet-features/src/home-commands.test.ts:330:1` |
| 16 | insert-drag rejects an insertion point inside its source before changing the workbook | `packages/sheet-features/src/home-commands.test.ts:356:1` |
| 17 | duplicate persisted rule identities reject clear and remove without partial writes | `packages/sheet-features/src/index.test.ts:935:1` |
| 18 | transposed paste keeps source-offset column-width mapping within its declared columns | `packages/sheet-features/src/index.test.ts:1018:1` |
| 19 | cross-sheet cut paste fails closed before changing either worksheet | `packages/sheet-features/src/index.test.ts:1176:1` |
| 20 | sheet commands: row insert/delete use StructuralTransform and preserve undo | `packages/sheet-features/src/index.test.ts:1719:1` |
| 21 | rejects a sparse row cell insertion that would discard the last cell | `packages/sheet-features/src/index.test.ts:2061:3` |
| 22 | rejects a sparse column cell insertion that would discard the last cell | `packages/sheet-features/src/index.test.ts:2061:3` |
| 23 | sorting uses resolved formula results, keeps stable ties, and replays/undoes as one permutation | `packages/sheet-features/src/m3-m4-data.test.ts:30:1` |
| 24 | rows.permuted rejects moving a manually hidden row before changing cells | `packages/sheet-features/src/m3-m4-data.test.ts:321:1` |
| 25 | rows.permuted rejects reordering rows inside an outline group before changing cells | `packages/sheet-features/src/m3-m4-data.test.ts:346:1` |
| 26 | sort and remove duplicates preserve formulas and use structural row deletion | `packages/sheet-features/src/m3-m4-data.test.ts:713:1` |
| 27 | Validation supports custom AST, formula-backed list, time/date, multi-select and non-blocking alerts | `packages/sheet-features/src/m3-m4-data.test.ts:1116:1` |
| 28 | acknowledges a pending commit already included in the hydrated snapshot without applying its patch twice | `packages/spreadsheet-app/src/application-collaboration.test.ts:276:3` |
| 29 | invalidates local undo after a committed range move without a canonical history transform | `packages/spreadsheet-app/src/application-collaboration.test.ts:528:3` |
| 30 | Fill Series keeps the complete selected seed range in the canonical planner | `packages/spreadsheet-app/src/application-core-editing.test.ts:31:3` |
| 31 | paste special values copies values without formulas | `packages/spreadsheet-app/src/application-core-editing.test.ts:119:3` |
| 32 | paste source theme adopts the clipboard theme as one undoable workbook state | `packages/spreadsheet-app/src/application-core-editing.test.ts:140:3` |
| 33 | returns a typed dispatch outcome and preserves state when permission rejects a paste | `packages/spreadsheet-app/src/application-core-editing.test.ts:157:3` |
| 34 | keeps paste pending until data-region materialization rejects and then preserves all state | `packages/spreadsheet-app/src/application-core-editing.test.ts:184:3` |
| 35 | cell insert down preserves the selected data and shifts the following band | `packages/spreadsheet-app/src/application-core-editing.test.ts:488:3` |
| 36 | Create Table plan registers a sheet table with explicit headers | `packages/spreadsheet-app/src/application-data-objects.test.ts:24:3` |
| 37 | auto-expands a table when a contiguous row is written below it | `packages/spreadsheet-app/src/application-data-objects.test.ts:51:3` |
| 38 | canonicalizes a command range to its sheet before matching lazy data regions | `packages/spreadsheet-app/src/application-data-tools.test.ts:217:3` |
| 39 | data dispatch keeps unique rows | `packages/spreadsheet-app/src/application-data-tools.test.ts:246:3` |
| 40 | saves locally without depending on a server role projection | `packages/spreadsheet-app/src/application-persistence.test.ts:32:3` |
| 41 | drillDownPivot creates a detail worksheet through the canonical command | `packages/spreadsheet-app/src/application-pivot.test.ts:288:3` |
| 42 | drillDownPivot grows the detail worksheet for more than the default row extent | `packages/spreadsheet-app/src/application-pivot.test.ts:302:3` |
| 43 | keeps an explicit block-backed worksheet source on the DataSource Pivot path | `packages/spreadsheet-app/src/application-pivot.test.ts:450:3` |
| 44 | loads DataSource Pivot field members only when a picker requests them | `packages/spreadsheet-app/src/application-pivot.test.ts:507:3` |
| 45 | creates a DataSource Pivot timeline from Query date fields | `packages/spreadsheet-app/src/application-pivot.test.ts:531:3` |
| 46 | refreshes one query without discarding another sheet data source reader | `packages/spreadsheet-app/src/application-query.test.ts:58:3` |
| 47 | sorts a block-backed query through AutoFilter without materializing cells | `packages/spreadsheet-app/src/application-query.test.ts:125:3` |
| 48 | does not persist an identity row order when a sorted block source only changes sort state | `packages/spreadsheet-app/src/application-query.test.ts:159:3` |
| 49 | keeps the current logical order for equal keys in a block-backed stable sort | `packages/spreadsheet-app/src/application-query.test.ts:212:3` |
| 50 | anchors the AutoFilter sort context to the full region after selecting a filter column | `packages/spreadsheet-app/src/application-query.test.ts:239:3` |
| 51 | builds the single client operation contract without server-owned fields | `packages/spreadsheet-app/src/collaboration/helpers.test.ts:31:3` |
| 52 | rejects malformed durable operation records and remote unknown mutations | `packages/spreadsheet-app/src/collaboration/helpers.test.ts:201:3` |
| 53 | rebases formula-rule parameters from their formula anchor sheet | `packages/spreadsheet-app/src/collaboration/ot-rebase.test.ts:220:1` |
| 54 | rebases every absolute coordinate in a pending paste snapshot | `packages/spreadsheet-app/src/collaboration/ot-rebase.test.ts:265:1` |
| 55 | fails closed when a pending cell restore carries unsupported formula-group metadata | `packages/spreadsheet-app/src/collaboration/ot-rebase.test.ts:488:1` |
| 56 | reads a canonical data-source Pivot source with stable field ids and source row paths | `packages/spreadsheet-app/src/features/pivot/block-source.test.ts:101:1` |
| 57 | returns explicit missing state instead of an empty source when a block is unavailable | `packages/spreadsheet-app/src/features/pivot/block-source.test.ts:135:1` |
| 58 | returns explicit error state for source identity mismatch and non-data-source Pivot sources | `packages/spreadsheet-app/src/features/pivot/block-source.test.ts:148:1` |
| 59 | drill-down creates a block-backed detail sheet and removes it through undo | `packages/spreadsheet-app/src/features/pivot/feature.test.ts:234:3` |
| 60 | drill-down resolves same-sheet joined rows by sourceId and recordId | `packages/spreadsheet-app/src/features/pivot/feature.test.ts:349:3` |
