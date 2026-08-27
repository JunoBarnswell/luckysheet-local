# Issue #304 PRD — Excel HOME / INSERT Canonical Ribbon

## 1. Product objective

Replace the mixed SpreadJS/project-specific HOME and INSERT ribbons with one Microsoft Excel-aligned product contract. The change is a clean break across UI schema, command ownership, model, rendering, interaction, history, collaboration, persistence, protocol, and OOXML. A visible surface is complete only when it reaches a real domain command or a typed unsupported boundary; no placeholder dialog, UI-only state, compatibility bridge, or silent fallback is allowed.

## 2. Evidence and visual contract

- Product inventory and behavior depth are defined by GitHub Issue #304, Microsoft Excel documentation, and the SpreadJS documentation cited by the issue.
- The user supplied the original Excel Chinese HOME image at 1783×141 and INSERT image at 1905×135 during implementation. They are the exact wide-ribbon pixel baselines for ordering, control hierarchy, split-button semantics, captions, launchers, focus states, spacing, typography, and dimensions.
- The existing Figma HOME assets are reusable only where their icon identity matches the Excel surface. They do not define INSERT geometry.
- Tailwind utilities and shared UI components remain the only presentation mechanism.

## 3. Canonical product information architecture

### HOME

1. Clipboard
2. Font
3. Alignment
4. Number
5. Styles
6. Cells
7. Editing

History is not a HOME group. Undo and Redo remain available through the application shell and shortcuts. Strikethrough remains a Format Cells/command capability and is not a direct HOME surface.

### INSERT

1. Tables
2. Illustrations
3. Controls
4. Charts
5. Sparklines
6. Filters
7. Links
8. Comments
9. Text
10. Symbols

TableSheet, GanttSheet, ReportSheet, Barcode, Camera, and floating legacy Form Controls leave the fixed Excel INSERT ribbon. Their existing domains may remain reachable from an explicit Extensions surface, but they cannot participate in the Excel schema.

## 4. Ownership and semantic chain

```text
RibbonSchema
  tab -> group -> surface -> menu item -> command id
        |
        v
Ribbon/shortcut/contextual tab/context menu
        |
        v
WorkbookSession.resolveCommandContext
        |
        v
CommandRuntime transaction
        |
        v
canonical model mutation
        |
        +-> history / undo / redo
        +-> collaboration operation
        +-> persistence snapshot
        +-> render projection
        `-> OOXML import/export
```

`RibbonSchema` owns ordering, appearance, menu membership, split primary action, launcher target, responsive priority, accessibility label, and command identity. React renderers may own transient DOM state such as an open menu or an editable font-size draft, but they do not own business selection state or invent menu membership.

## 5. State ownership

| State | Owner | Consumers |
|---|---|---|
| Selection and mixed formatting | `WorkbookSession` selection aggregate | HOME controls, Format Cells |
| Ribbon structure | `RibbonSchema` | HOME/INSERT/layout renderer/tests |
| Menu open/focus | shared UI components | renderer only |
| Format Painter mode | canonical interaction state machine | Ribbon, canvas, Esc handler |
| Drawing insertion mode | drawing interaction coordinator | Ribbon, canvas, history |
| Chart data binding | `ChartDrawingPayload.source` and bindings | renderer, editor, OOXML |
| Unsupported Office-host object | typed object metadata + opaque OOXML relationship | UI inspector, export |

The same state is never duplicated in a component and a store. Failed prerequisites return a typed observable error and must not mutate model, history, collaboration, or UI success state.

## 6. Chart clean break

`DrawingKind` contains only `chart`; `data-chart` is removed. The canonical chart payload owns:

- `source`: worksheet ranges, Pivot, DataManager table, or ReportSheet range;
- `family` and `subtype`: column/bar, line/area, pie/doughnut, hierarchy, statistical, scatter/bubble, waterfall/funnel/stock/surface/radar, combo, map;
- series/bindings, axes, legend, labels, plot/chart styles, stacking, secondary axis, and hidden-data policy.

DataManager field bindings are a chart-source concern, not a drawing kind. PivotChart is a chart with a Pivot source and retains filter/slicer/timeline linkage. Recommended Charts produces deterministic recommendations from typed selection statistics and inserts the selected canonical chart. Runtime snapshots accept only the canonical chart contract.

## 7. Drawing and Office object contract

- Shapes use one preset identity shared by gallery, renderer, hit-test, serialization, and OOXML preset geometry. Drawing mode and lock mode are explicit transient interaction states.
- Checkbox writes a real boolean cell value and cell control formatting; it never creates a floating form-control drawing.
- Picture placement distinguishes in-cell and floating placement and can convert through the same asset reference.
- Icons, 3D Models, SmartArt, Screenshot, Text Box, WordArt, Signature Line, Embedded Object, and Equation use typed drawing/object payloads when editable semantics are implemented.
- When an Office-host operation cannot execute in the browser (for example screen clipping, digital signing, or opening an OLE server), its command returns `UNSUPPORTED_FEATURE` with feature id, affected target, reason, and recovery action. Existing opaque OOXML parts and relationships are retained for round trip; the UI never reports insertion success without a canonical mutation.
- Symbol inserts a Unicode scalar into the active text-edit target and fails closed when no text target exists.

## 8. HOME behavior contract

- Paste is a real split button. The menu covers All, Formulas, Values, Formats, Comments/Notes, Validation, Source Theme, All Except Borders, Column Widths, Formulas + Number Formats, Values + Number Formats, Skip Blanks, Transpose, Paste Link, and Add/Subtract/Multiply/Divide.
- Format Painter supports one-shot, double-click lock, and Esc cancellation.
- Font, Alignment, and Number dialog launchers open the exact Format Cells section.
- Font size follows the Excel step table rather than arithmetic one-point increments.
- Underline, Borders, Fill, Font Color, Phonetic Guide, Orientation, Merge, Number formats, Conditional Formatting, Format as Table, Cell Styles, Insert/Delete/Format, AutoSum, Fill, Clear, Sort & Filter, and Find & Select expose their complete menu contract.
- Each command validates sheet, selection, permissions, and source data before opening a transaction.

## 9. INSERT behavior contract

- Tables: PivotTable, Recommended PivotTables, Table, Forms.
- Illustrations: Pictures, Shapes, Icons, 3D Models, SmartArt, Screenshot.
- Controls: Boolean cell Checkbox.
- Charts: Recommended Charts plus Excel family menus and PivotChart.
- Sparklines: Line, Column, Win/Loss.
- Filters: Slicer and Timeline, reusing the Pivot contextual command owners.
- Links: Hyperlink.
- Comments: threaded comment owner with reply, mention, resolve, and reopen.
- Text: Text Box, Header & Footer, WordArt, Signature Line, Object.
- Symbols: Equation and Symbol.

Recommended surfaces never return invented sample data. An empty or unsuitable selection is a typed rejection with a recovery instruction.

## 10. Responsive and accessibility contract

- Wide mode preserves group order and never moves a command to a different business group.
- Compact/narrow modes collapse a complete group to a schema-derived menu/gallery; they do not reorder or split the group.
- Split buttons expose separate primary and dropdown focus targets.
- Every surface and menu item has an accessible name, keyboard focus state, disabled reason, and deterministic focus order.
- Hover, pressed, selected, disabled, and focus-visible states use shared Tailwind tokens.

## 11. Persistence, migration, and OOXML

- This clean break does not retain or infer the obsolete `data-chart` contract. Hydrate/import accepts canonical `chart.source` only; a legacy snapshot is rejected with `MIGRATION_REQUIRED` and must be converted by an explicit offline migration owned outside runtime.
- Runtime validators, mutation descriptors, collaboration operations, and protocol DTOs contain no legacy `data-chart` aliases.
- Unknown OOXML parts, nodes, extensions, macros, embedded objects, signatures, and relationships remain opaque and round-trip unchanged unless an explicit supported edit owns their replacement.
- Supported edits update the canonical model and the corresponding OOXML part in one export transaction.

## 12. Failure model

Every rejection includes `code`, `feature`, `reason`, `affectedObject`, and `recovery`. Required codes include `UNSUPPORTED_FEATURE`, `INVALID_SELECTION`, `PERMISSION_DENIED`, `INVALID_CHART_SOURCE`, `MIGRATION_REQUIRED`, and `OOXML_FIDELITY_VIOLATION`. No caller may catch these errors to return an empty result or success toast.

## 13. Acceptance and verification

1. Static architecture review proves one Ribbon schema, one Chart model, one command owner per surface, and no bridge/double-write/fallback path.
2. Success and rejection tests cover every fail-close or clean-break contract.
3. Unit, typecheck, boundaries, contracts, acceptance matrix, and production build pass.
4. Menu-level Playwright checks cover every dropdown and each Format Cells launcher.
5. XLSX corpus round trips tables, charts, Pivot, slicer/timeline, pictures, shapes, comments, hyperlinks, header/footer, and opaque embedded objects.
6. In-app browser checks page identity, meaningful DOM, framework overlay, console, network, HOME/INSERT interactions, persistence behavior, keyboard focus, and screenshots.
7. Exact 1783×141 HOME and 1905×135 INSERT captures are compared directly with the supplied source pixels. The Pixel Difference Report records component-level mismatch and no substitute screenshot may be declared equivalent.

## 14. Delivery and rollback

The PR is delivered from `codex/issue-304-excel-ribbon-parity`. Rollback is a single PR revert because the runtime migration is one-way at an explicit snapshot boundary and no dual-write path exists. If persisted legacy data cannot be migrated unambiguously, merge is blocked rather than reintroducing a compatibility reader.
