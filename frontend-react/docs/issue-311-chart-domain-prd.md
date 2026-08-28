# Issue #311 — Canonical Excel Chart Domain PRD

## Product contract

Chart, PivotChart, and Sparkline are spreadsheet objects, not Canvas decorations.
Every insertion or edit resolves through one ChartDomain chain:

`ChartSource -> ChartBindingModel -> ChartSpec/Presentation -> ChartLayout -> Renderer/HitTest and native OOXML`

Worksheet ranges, Table, Report, and Pivot are source variants of the same
domain. The UI only dispatches typed commands. The renderer consumes layout
geometry and never reads worksheet cells or invents data semantics.

## Owned model

- `ChartSource` owns worksheet ranges, table structured references, report
  ranges, Pivot identity, orientation, and dynamic/non-contiguous identity.
- `ChartDrawingPayload` owns family/subtype, typed series, axes, elements,
  statistical options, map state, and native identity. `DrawingObject` owns
  placement only.
- Scatter uses numeric X/Y ranges. Bubble uses numeric X/Y/Size ranges.
  Stock uses explicit Open/High/Low/Close/Volume roles. Combo stores the type
  and axis group on every series.
- Empty, hidden, and filtered cells are resolved before layout. A missing
  source or unresolved geographic entity is an observable invalid or
  unsupported state.
- Sparkline groups own shared type, scale, orientation, hidden/empty behavior,
  and presentation options; a member does not keep a conflicting second copy.

## UI contract

Insert exposes all supported Excel families and subtypes. Recommended Charts
uses the current selection data to build candidates and real data previews.
Chart Design/Format and Select Data dispatch domain commands for elements,
series order, axes, labels, data table, trendlines, error bars, and formats.
Canvas clicks resolve chart area, plot area, title, series, and point identity
before the generic drawing gesture.

## Native I/O contract

Owned charts are emitted as native DrawingML chart parts with worksheet
anchors, relationships, plot family, series references, axes, elements, and
blank-cell behavior. Imported owned parts project into the canonical model;
unknown extensions are retained. Legal chart identities without complete
editor ownership remain `preserved-native` and are never converted to Column.

## Clean-break removals

The obsolete `data-chart` contract, category-index Scatter X, Y-derived Bubble
size, fixed Combo ordering, static Recommended Chart icon previews, and the
old Pivot-only native writer are not runtime fallback paths. Unsupported
native behavior fails with a typed observable reason or remains in the native
package preservation domain.

## Acceptance matrix

1. Each family/subtype creates a typed payload and deterministic layout.
2. Scatter/Bubble/Stock/statistical/Waterfall/Surface/Radar/Combo data fixtures
   prove their mathematical input roles.
3. Element selection, formatting, history, copy/paste, and remote replay use
   the same command payload.
4. Recommended previews change with the selected data.
5. Pivot state drives PivotChart; forbidden PivotChart families are rejected.
6. Sparkline group/member options round-trip with hidden/empty/date/scale
   semantics.
7. Owned Chart and PivotChart round-trip through native OOXML; unsupported
   native parts are preserved and reported.

## Acceptance evidence — 2026-08-28

- Source workbook: `OCR结果.xlsx`; OOXML contains 10 base parts and no existing
  `xl/charts`, `xl/drawings`, or Sparkline parts. The real OCR data range used
  for chart acceptance was `Sheet1!I1:N20`, whose numeric-looking text values
  are resolved by the canonical chart numeric resolver without changing the
  source cell types.
- Runtime/layout matrix: 80 subtypes checked through the imported workbook;
  78 returned `ready` across Column, Bar, Line, Area, Pie, Doughnut, Scatter,
  Bubble, Treemap, Sunburst, Histogram, Pareto, Box & Whisker, Waterfall,
  Funnel, Stock, Surface, Radar, and Combo. Filled Map and Region Map return
  `UNSUPPORTED_FEATURE` because no authoritative geographic geometry provider
  is configured.
- Native OOXML matrix: the same 80 subtypes checked through the native writer;
  78 emitted one structurally complete `xl/charts/react-chart-*` part. The two
  map subtypes fail closed before emitting a fake chart.
- UI acceptance: imported OCR data opens in the editor; Insert Chart gallery,
  all family/subtype labels, Chart Design/Format, Select Data, data labels,
  and Sparkline Design are reachable. Recommended Charts produced real
  Scatter and Bubble previews from OCR numeric strings. Console and failed
  network request checks were empty in terminal-browser fallback validation.
- Desktop Excel round-trip is not claimed in this environment; it remains
  `Blocked` until native Excel is available.
