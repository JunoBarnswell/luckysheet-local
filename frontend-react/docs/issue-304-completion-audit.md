# Issue #304 Completion Audit

Status: in progress. This file is the requirement-to-evidence ledger for the clean-break implementation. A green build is not completion.

## Evidence grades

- `PROVEN`: current source plus an applicable test/runtime artifact proves the requirement.
- `PARTIAL`: the canonical owner exists but behavior depth or verification is incomplete.
- `MISSING`: no real domain implementation exists, or the current surface only rejects/opens a generic substitute.
- `BLOCKED`: implementation exists but required external acceptance evidence cannot currently be collected.

## Architecture

| Requirement | Status | Current evidence | Required closure |
|---|---|---|---|
| One HOME/INSERT Ribbon schema | PROVEN | `ui-command-catalog/index.ts` owns groups, surfaces and layout; renderer consumes it | Browser DOM readback |
| No HOME History group | PROVEN | History group/surfaces removed; undo/redo moved out of HOME | Screenshot comparison |
| No direct Strikethrough | PROVEN | direct surface removed; Format Cells font tab retains the property | Launcher interaction |
| One Chart domain | PROVEN | `data-chart` kind/DTO/panel/commands removed; `ChartSource` owns worksheet/Pivot/table/report data; structured table/report resolver is shared by runtime and canvas | Full chart-family behavior and OOXML corpus |
| No runtime legacy Chart reader | PROVEN | snapshot rejects legacy `data-chart` with `MIGRATION_REQUIRED` | Rejection test at hydrate boundary |
| One command/history/collaboration/persistence chain | PARTIAL | canonical insert/update commands and mutation paths exist | Per-surface evidence matrix and remote replay coverage |
| One Shape preset owner | PARTIAL | `SHAPE_DRAWING_PRESETS` is now the single model/gallery identity registry; renderer currently implements seven presets | Extend the registry only with renderer/hit-test/OOXML-backed DrawingML presets |

## HOME

| Surface/contract | Status | Current evidence | Required closure |
|---|---|---|---|
| Paste split + Paste Special depth | PROVEN | split surface opens the canonical dialog; 20 typed modes and arithmetic operations exist | Browser menu/transaction test |
| Cut/Copy | PROVEN | canonical clipboard/session path | Browser/system clipboard proof |
| Format Painter single/locked/Esc | PARTIAL | single and double-click modes exist; domain tests cover formatting | Browser Esc and lock-loop interaction |
| Font family/size/step | PROVEN | mixed selection aggregate and Excel font-size ladder | Browser mixed state |
| Underline/Borders/Fill/Font Color | PARTIAL | canonical commands and galleries exist | Double underline, border color/line drawing, recent colors |
| Phonetic Guide | PROVEN | canonical cell metadata, one cell.set mutation chain, dialog, render measurement/guide and OOXML rPh/phoneticPr round trip | Browser success-path edit with a non-empty East Asian selection |
| Font/Alignment/Number launchers | PROVEN | each launcher carries an explicit Format Cells tab in session state | Browser target-tab assertion |
| Alignment/Orientation/Wrap/Merge | PARTIAL | canonical style/merge commands exist | complete angle UI and all menu tests |
| Number gallery/accounting/precision | PARTIAL | number-format model and precision transformer exist | complete locale/accounting dropdown and More Formats |
| Conditional Formatting | PARTIAL | rule domain and manager panel exist | complete Excel rule galleries/priority UI |
| Format as Table | PROVEN | creates a real Sheet Table | style gallery/browser proof |
| Cell Styles | PARTIAL | built-in styles exist | New/Merge Styles behavior |
| Cells Insert/Delete/Format | PARTIAL | row/column/cell/dimension commands exist | sheet operations, move/copy, tab color, protect sheet menu |
| AutoSum | PROVEN | typed adjacent-region inference and canonical formula mutation | menu browser proof |
| Fill | PARTIAL | directional/series planner exists | Justify, Flash Fill, date/custom-list depth |
| Clear | PARTIAL | all required clear families exist | Remove Hyperlinks distinct action |
| Sort & Filter | PARTIAL | canonical sort/filter exists | complete Home menu membership and Advanced/Reapply actions |
| Find & Select | PARTIAL | Find/Replace/Go To/Go To Special and Selection Pane domains exist | complete menu membership and select-object browser flow |

## INSERT

| Surface/contract | Status | Current evidence | Required closure |
|---|---|---|---|
| PivotTable | PROVEN | real create dialog/task/field pane/domain | Browser creation and persistence |
| Recommended PivotTables | PROVEN | deterministic field-type analyzer, preview dialog, selected candidate creation through Pivot command/history | Browser success-path creation |
| Table | PROVEN | real Sheet Table creation | browser Ctrl+T parity |
| Forms | MISSING | current surface returns `UNSUPPORTED_FEATURE` | form definition, response table association and persistence |
| Pictures | PARTIAL | asset store, in-cell/floating conversion, crop/effects exist | split sources and browser insertion flow |
| Shapes | MISSING | only seven renderer-backed identities | full categorized preset registry and drawing mode |
| Icons | MISSING | current surface returns `UNSUPPORTED_FEATURE` | searchable asset library and canonical image/vector object |
| 3D Models | MISSING | current surface returns `UNSUPPORTED_FEATURE` | asset payload, view/scene state, renderer and OOXML preservation/edit boundary |
| SmartArt | MISSING | current surface returns `UNSUPPORTED_FEATURE` | typed layouts/text pane/style and OOXML graph |
| Screenshot | MISSING | current surface returns `UNSUPPORTED_FEATURE` | Screen Capture API command and picture insertion |
| Checkbox | PROVEN | canonical boolean cell editor/toggle, not floating control | Browser click/Space proof |
| Recommended Charts | PROVEN | deterministic selection analyzer, preview dialog, selected candidate insertion through canonical Chart command | Browser success-path creation with populated selection |
| Chart families | PARTIAL | model names all requested families | subtype contract and real renderer/OOXML for every family |
| PivotChart | PROVEN | canonical Chart with Pivot source and shared Pivot command | browser filter linkage/XLSX proof |
| Sparklines | PROVEN | line/column/win-loss canonical commands | three fixed-entry browser proof |
| Slicer/Timeline | PROVEN | fixed and contextual entries share command owners | browser fixed-entry proof |
| Link | PROVEN | typed URL/file/workbook/email target | browser Ctrl+K parity |
| Threaded Comment | PARTIAL | thread/reply/resolve author/time domain exists | mention and reopen behavior |
| Text Box | PROVEN | placement/edit/text-frame mutation chain | OOXML round trip |
| Header & Footer | PARTIAL | print model/panel exists | section tokens and OOXML round trip |
| WordArt | MISSING | current surface returns `UNSUPPORTED_FEATURE` | typed transform/style text object |
| Signature Line | MISSING | current surface returns `UNSUPPORTED_FEATURE` | metadata/status payload and opaque signature fidelity |
| Object/OLE | MISSING | current surface returns `UNSUPPORTED_FEATURE` | embedded/linked file payload and relationships |
| Equation | MISSING | current surface returns `UNSUPPORTED_FEATURE` | Office Math model/editor/renderer/OOXML |
| Symbol | PROVEN | Unicode picker/recent list and active cell/text-box target mutation with fail-close target validation | Browser success-path insertion into an edited cell |
| Extension items removed from fixed INSERT | PROVEN | TableSheet/Gantt/Report/Barcode/Camera/Form Controls no longer have fixed INSERT surfaces | Browser DOM proof |

## Verification and delivery

| Gate | Status | Evidence |
|---|---|---|
| Typecheck | PROVEN | `npm run typecheck` passed |
| Unit suite | PROVEN | 737/737 passed |
| Boundaries/contracts/acceptance/build | PROVEN | all passed; only existing Vite chunk warnings |
| In-app browser | PROVEN | in-app browser captured HOME/INSERT at 1783/1905 viewports; DOM, dialog, disabled-state and console checks passed; captures are recorded in `design-qa.md` |
| Pixel Difference Report | PROVEN | component-level report and final captures recorded in `design-qa.md`; no P0/P1 mismatch observed |
| Native Excel corpus | MISSING | desktop Excel round-trip/repair-prompt evidence not run |
| PR/checks/merge | MISSING | no PR exists; Issue remains open |

The issue cannot be closed while any `MISSING`, `PARTIAL`, or `BLOCKED` row remains.
