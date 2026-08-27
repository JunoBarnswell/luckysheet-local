# Figma Home Ribbon Pixel-Perfect PRD

## Product objective

Reproduce Figma file `B3SP4FCTDr4IMlvq1CKc6M`, node `3:28` as the Home Ribbon visual contract while preserving the existing canonical command, selection-state, menu, permission, persistence, collaboration, and workbook runtime paths.

The target is visual and functional parity. No command may be mocked, duplicated, hardcoded around the catalog, or redirected through a compatibility path.

## Figma source contract

- Root frame: `toolbar-container`, `1500 × 101`, horizontal Auto Layout.
- Root alignment: vertically centered.
- Root spacing: `8px` between every group/divider child.
- Root padding: `12px` horizontal, `4px` vertical.
- Root behavior: width fills its host; `1500px` is the minimum complete layout width; horizontal scrolling preserves the layout below that width.
- Group structure: `70px` control area, `4px` control-to-caption spacing, `10px/12px` caption, `2px` bottom padding.
- Divider: `1 × 76px`, centered vertically.
- Group widths at the 1500px design frame:
  - history `64px`
  - clipboard `118px`
  - font `205px`
  - alignment `211px`
  - number `120px`
  - styles `266px`
  - cells `158px`
  - editing `222px`

## Layer and component tree

```text
RibbonShell
└── RibbonLayoutRenderer (Home visual owner)
    ├── HistoryGroup
    │   └── Undo / Redo
    ├── ClipboardGroup
    │   ├── Paste
    │   └── Cut / Copy / FormatPainter
    ├── FontGroup
    │   ├── FontFamily / FontSize / Increase / Decrease
    │   └── Bold / Italic / Underline / Strike / Border / Fill / FontColor
    ├── AlignmentGroup
    │   ├── VerticalAlignment / HorizontalAlignment
    │   ├── WrapText / MergeCenter
    │   └── Orientation
    ├── NumberGroup
    │   ├── NumberFormat
    │   └── Percent / Comma / DecimalIncrease / DecimalDecrease
    ├── StylesGroup
    │   └── ConditionalFormatting / FormatAsTable / CellStyles
    ├── CellsGroup
    │   └── Insert / Delete / Format
    └── EditingGroup
        ├── AutoSum / Fill / Clear
        └── SortAndFilter / FindAndSelect
```

`RibbonLayoutRenderer` continues to consume `RIBBON_LAYOUT_SPECS`; `HomeRibbon` continues to resolve every surface from `getRibbonSurfaces`; `CatalogButton` continues to execute `buildRibbonCommand`. The change does not introduce a second command or state chain.

## Design tokens

### Color

- background/surface: `#ffffff`
- primary text: `#333333`
- caption/muted text: `#999999`
- control border: `#d0d4dc`
- hover surface: existing shared hover feedback, constrained to the same geometry

### Typography

- family: `Noto Sans SC` variable font
- command label: `12px`, regular, normal line-height (`14px` measured in Figma)
- compact/input label: `11px`, regular, `13px` measured line box
- group caption: `10px`, regular, `12px` measured line box
- font scale affordance: `11px`, bold
- letter spacing: default (`0`)

### Spacing and geometry

- base steps: `2 / 4 / 6 / 8 / 12px`
- radius: `4px`
- small icon: `14px`
- regular icon: `16px`
- history icon: `20px`
- large icon: `32px`
- compact row height: `24px`
- large button height: `72px`

### Effects

- no shadow or blur in the source node
- opacity `100%`; disabled state remains owned by the shared Button contract

## Responsive and failure behavior

- At widths below `1500px`, the Home Ribbon scrolls horizontally; groups and controls are not compressed or truncated.
- Menus remain keyboard reachable and preserve their current command enablement rules.
- Missing committed SVG or font assets are a build/acceptance failure; runtime fallback icons are not permitted for the Figma Home surface.
- Unsupported or disabled commands retain the existing observable disabled state.

## Acceptance plan

1. Structural iteration: match the layer tree and group geometry.
2. Spacing iteration: match root/group gaps, padding, controls, dividers, and scroll behavior.
3. Typography iteration: load and apply Noto Sans SC; match sizes, weights, line boxes, and baselines.
4. Visual polish iteration: use committed Figma SVGs and match border/radius/state visuals.
5. Capture the in-app browser at the reference scale and report width, height, padding, baseline, icon-position, and visual-state differences.
6. Run typecheck, focused layout/localization/icon tests, unit tests, build, desktop packaging checks, and in-app browser console/network acceptance.
