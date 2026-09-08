//! Canonical worksheet geometry.  This crate owns model-to-screen coordinates;
//! callers must use the returned pane map for drawing, hit testing and editing.
use kernel_core::{CellAddress, KernelError, KernelResult, MAX_COLUMNS, MAX_ROWS};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};

pub const ROW_HEADER_WIDTH: f64 = 39.0;
pub const COL_HEADER_HEIGHT: f64 = 20.0;

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Point {
    pub x: f64,
    pub y: f64,
}
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Rect {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
}
#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Size {
    pub width: f64,
    pub height: f64,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CellRange {
    pub start_row: u32,
    pub end_row: u32,
    pub start_column: u32,
    pub end_column: u32,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub enum PaneId {
    TopLeft,
    TopRight,
    BottomLeft,
    Main,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct RenderPane {
    pub id: PaneId,
    pub screen_rect: Rect,
    pub content_origin: Point,
    pub visible_range: Option<CellRange>,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PaneMap {
    pub panes: Vec<RenderPane>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Viewport {
    pub width: f64,
    pub height: f64,
    pub scroll_x: f64,
    pub scroll_y: f64,
    pub device_pixel_ratio: f64,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PaneLayout {
    pub kind: PaneKind,
    pub x_split: f64,
    pub y_split: f64,
    pub start_row: u32,
    pub start_column: u32,
    #[serde(default)]
    pub active_pane: Option<PaneId>,
    pub state: PaneState,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PaneKind {
    None,
    Frozen,
    Split,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum PaneState {
    Frozen,
    FrozenSplit,
    Split,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct GeometryRequest {
    pub sheet_id: String,
    pub row_count: u32,
    pub column_count: u32,
    pub default_row_height_px: f64,
    pub default_column_width_px: f64,
    #[serde(default)]
    pub row_heights_px: BTreeMap<u32, f64>,
    #[serde(default)]
    pub column_widths_px: BTreeMap<u32, f64>,
    #[serde(default)]
    pub hidden_rows: BTreeSet<u32>,
    #[serde(default)]
    pub hidden_columns: BTreeSet<u32>,
    #[serde(default = "one")]
    pub zoom: f64,
    pub viewport: Viewport,
    #[serde(default)]
    pub pane: Option<PaneLayout>,
    #[serde(default)]
    pub header_offset: Option<Point>,
}
fn one() -> f64 {
    1.0
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct GeometryResponse {
    pub pane_map: PaneMap,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct HitTestResponse {
    pub address: Option<CellAddress>,
    pub pane: Option<PaneId>,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct HeaderRectResponse {
    pub rect: Rect,
    pub axis: HeaderAxis,
    pub index: Option<u32>,
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum HeaderAxis {
    Corner,
    Row,
    Column,
}

#[derive(Clone)]
struct Axis {
    count: u32,
    default: f64,
    zoom: f64,
    overrides: BTreeMap<u32, f64>,
    hidden: BTreeSet<u32>,
    adjustments: Vec<(u32, f64)>,
    prefix: Vec<f64>,
}
impl Axis {
    fn new(
        count: u32,
        default: f64,
        zoom: f64,
        overrides: BTreeMap<u32, f64>,
        hidden: BTreeSet<u32>,
        axis: &'static str,
    ) -> KernelResult<Self> {
        if count == 0 || count > if axis == "row" { MAX_ROWS } else { MAX_COLUMNS } {
            return Err(KernelError::new(
                "GEOMETRY_DIMENSION_INVALID",
                "Dimension count is outside the worksheet bounds",
            )
            .at(axis));
        }
        if !default.is_finite() || default < 0.0 || !zoom.is_finite() || zoom <= 0.0 {
            return Err(KernelError::new(
                "GEOMETRY_METRIC_INVALID",
                "Default dimensions and zoom must be finite and non-negative",
            )
            .at(axis));
        }
        let mut adjustments = Vec::new();
        let mut keys = BTreeSet::new();
        for (&i, &v) in &overrides {
            if i >= count || !v.is_finite() || v < 0.0 {
                return Err(KernelError::new(
                    "GEOMETRY_OVERRIDE_INVALID",
                    "Dimension override is outside the worksheet or invalid",
                )
                .at(axis));
            }
            keys.insert(i);
        }
        for &i in &hidden {
            if i >= count {
                return Err(KernelError::new(
                    "GEOMETRY_HIDDEN_INDEX_INVALID",
                    "Hidden dimension index is outside the worksheet",
                )
                .at(axis));
            }
            keys.insert(i);
        }
        for i in keys {
            let value = if hidden.contains(&i) {
                0.0
            } else {
                overrides.get(&i).copied().unwrap_or(default) * zoom
            };
            adjustments.push((i, value - default * zoom));
        }
        let mut prefix = vec![0.0];
        for (_, d) in &adjustments {
            prefix.push(prefix.last().unwrap() + d);
        }
        Ok(Self {
            count,
            default,
            zoom,
            overrides,
            hidden,
            adjustments,
            prefix,
        })
    }
    fn total_before(&self, i: u32) -> f64 {
        let i = i.min(self.count);
        let n = self.adjustments.partition_point(|(at, _)| *at < i);
        self.default * self.zoom * i as f64 + self.prefix[n]
    }
    fn extent(&self, start: u32, end: u32) -> f64 {
        if end < start {
            0.0
        } else {
            self.total_before(end.saturating_add(1)) - self.total_before(start)
        }
    }
    fn size(&self, i: u32) -> f64 {
        self.extent(i, i)
    }
    fn index_at(&self, coordinate: f64) -> Option<u32> {
        if !coordinate.is_finite()
            || coordinate < 0.0
            || coordinate >= self.total_before(self.count)
        {
            return None;
        }
        let mut lo = 0;
        let mut hi = self.count;
        while lo < hi {
            let m = lo + (hi - lo) / 2;
            if self.total_before(m + 1) <= coordinate {
                lo = m + 1;
            } else {
                hi = m;
            }
        }
        if lo < self.count && self.size(lo) > 0.0 {
            Some(lo)
        } else {
            None
        }
    }
    fn range(&self, start: f64, extent: f64) -> Option<(u32, u32)> {
        if extent <= 0.0 {
            return None;
        }
        let first = self.index_at(start.max(0.0))?;
        let last_coord = (start + extent)
            .min(self.total_before(self.count))
            .max(start);
        // Pane extents are half-open. Subtract one representable float rather
        // than f64::EPSILON: the latter is smaller than one ULP once worksheet
        // coordinates grow beyond 1px and can leave an exact total unchanged.
        let exclusive_end = if last_coord > 0.0 {
            f64::from_bits(last_coord.to_bits() - 1)
        } else {
            last_coord
        };
        let mut last = self.index_at(exclusive_end.max(start))?;
        if last < first {
            last = first;
        }
        Some((first, last))
    }
}

struct Geometry {
    rows: Axis,
    columns: Axis,
    viewport: Viewport,
    origin: Point,
    sheet_id: String,
}
impl Geometry {
    fn from(req: &GeometryRequest) -> KernelResult<Self> {
        if req.sheet_id.is_empty() {
            return Err(KernelError::new(
                "GEOMETRY_SHEET_INVALID",
                "Sheet id must not be empty",
            ));
        }
        let v = &req.viewport;
        if ![
            v.width,
            v.height,
            v.scroll_x,
            v.scroll_y,
            v.device_pixel_ratio,
        ]
        .iter()
        .all(|x| x.is_finite())
            || v.width < 0.0
            || v.height < 0.0
            || v.scroll_x < 0.0
            || v.scroll_y < 0.0
            || v.device_pixel_ratio <= 0.0
        {
            return Err(KernelError::new(
                "GEOMETRY_VIEWPORT_INVALID",
                "Viewport values are invalid",
            ));
        }
        let zoom = req.zoom;
        let origin = req.header_offset.unwrap_or(Point {
            x: ROW_HEADER_WIDTH,
            y: COL_HEADER_HEIGHT,
        });
        if !origin.x.is_finite() || !origin.y.is_finite() || origin.x < 0.0 || origin.y < 0.0 {
            return Err(KernelError::new(
                "GEOMETRY_HEADER_INVALID",
                "Header offset is invalid",
            ));
        }
        Ok(Self {
            rows: Axis::new(
                req.row_count,
                req.default_row_height_px,
                zoom,
                req.row_heights_px.clone(),
                req.hidden_rows.clone(),
                "row",
            )?,
            columns: Axis::new(
                req.column_count,
                req.default_column_width_px,
                zoom,
                req.column_widths_px.clone(),
                req.hidden_columns.clone(),
                "column",
            )?,
            viewport: v.clone(),
            origin,
            sheet_id: req.sheet_id.clone(),
        })
    }
    fn visible(&self, origin: Point, rect: Rect) -> Option<CellRange> {
        let (sc, ec) = self.columns.range(origin.x, rect.width)?;
        let (sr, er) = self.rows.range(origin.y, rect.height)?;
        Some(CellRange {
            start_row: sr,
            end_row: er,
            start_column: sc,
            end_column: ec,
        })
    }
    fn pane_map(&self, pane: Option<&PaneLayout>) -> KernelResult<PaneMap> {
        let gw = (self.viewport.width - self.origin.x).max(0.0);
        let gh = (self.viewport.height - self.origin.y).max(0.0);
        if let Some(pane) = pane {
            validate_pane_split(pane)?;
        }
        let p = pane.filter(|p| p.kind != PaneKind::None && (p.x_split > 0.0 || p.y_split > 0.0));
        if p.is_none() {
            let rect = Rect {
                x: self.origin.x,
                y: self.origin.y,
                width: gw,
                height: gh,
            };
            return Ok(PaneMap {
                panes: vec![RenderPane {
                    id: PaneId::Main,
                    screen_rect: rect,
                    content_origin: Point {
                        x: self.viewport.scroll_x,
                        y: self.viewport.scroll_y,
                    },
                    visible_range: self.visible(
                        Point {
                            x: self.viewport.scroll_x,
                            y: self.viewport.scroll_y,
                        },
                        rect,
                    ),
                }],
            });
        }
        let p = p.unwrap();
        let frozen = p.kind == PaneKind::Frozen;
        let fx = if frozen {
            freeze_split_count(p.x_split, self.columns.count, "xSplit")?
        } else {
            0
        };
        let fy = if frozen {
            freeze_split_count(p.y_split, self.rows.count, "ySplit")?
        } else {
            0
        };
        // A zero split is a real single-axis freeze boundary.  Calling
        // `extent(0, 0)` for it would incorrectly reserve the first row or
        // column even though that axis has no frozen cells.
        let left = if fx == 0 {
            0.0
        } else {
            self.columns.extent(0, fx - 1).min(gw)
        };
        let top = if fy == 0 {
            0.0
        } else {
            self.rows.extent(0, fy - 1).min(gh)
        };
        let sx = if frozen {
            // The viewport scroll is already expressed in worksheet content
            // coordinates.  When a saved frozen start cell seeded it, adding
            // the frozen extent again skips that same extent twice.  Keep the
            // main pane at or beyond the frozen boundary while preserving a
            // normal zero-scroll freeze.
            self.columns.total_before(fx).max(self.viewport.scroll_x)
        } else {
            // OOXML split positions are twentieths of a point.  Convert at
            // the geometry boundary to CSS pixels (96 CSS px per inch).
            p.x_split * (96.0 / (72.0 * 20.0))
        };
        let sy = if frozen {
            self.rows.total_before(fy).max(self.viewport.scroll_y)
        } else {
            p.y_split * (96.0 / (72.0 * 20.0))
        };
        let entries = [
            (
                PaneId::TopLeft,
                Rect {
                    x: self.origin.x,
                    y: self.origin.y,
                    width: left,
                    height: top,
                },
                Point { x: 0.0, y: 0.0 },
            ),
            (
                PaneId::TopRight,
                Rect {
                    x: self.origin.x + left,
                    y: self.origin.y,
                    width: gw - left,
                    height: top,
                },
                Point { x: sx, y: 0.0 },
            ),
            (
                PaneId::BottomLeft,
                Rect {
                    x: self.origin.x,
                    y: self.origin.y + top,
                    width: left,
                    height: gh - top,
                },
                Point { x: 0.0, y: sy },
            ),
            (
                PaneId::Main,
                Rect {
                    x: self.origin.x + left,
                    y: self.origin.y + top,
                    width: gw - left,
                    height: gh - top,
                },
                Point { x: sx, y: sy },
            ),
        ];
        let mut panes = Vec::new();
        for (id, rect, content) in entries {
            if rect.width <= 0.0 || rect.height <= 0.0 {
                continue;
            }
            let mut range = self.visible(content, rect);
            if frozen {
                range = clamp_range(id, range, fy, fx, self.rows.count, self.columns.count);
            }
            panes.push(RenderPane {
                id,
                screen_rect: rect,
                content_origin: content,
                visible_range: range,
            });
        }
        for a in 0..panes.len() {
            for b in (a + 1)..panes.len() {
                if overlap(panes[a].visible_range, panes[b].visible_range) {
                    return Err(KernelError::new(
                        "GEOMETRY_PANE_OVERLAP",
                        "Frozen pane visible ranges overlap; pane ownership is ambiguous",
                    ));
                }
            }
        }
        Ok(PaneMap { panes })
    }
}

fn validate_pane_split(pane: &PaneLayout) -> KernelResult<()> {
    if !pane.x_split.is_finite()
        || !pane.y_split.is_finite()
        || pane.x_split < 0.0
        || pane.y_split < 0.0
    {
        return Err(KernelError::new(
            "GEOMETRY_PANE_INVALID",
            "Pane split positions must be finite and non-negative",
        ));
    }
    if pane.kind == PaneKind::Frozen && (pane.x_split.fract() != 0.0 || pane.y_split.fract() != 0.0)
    {
        return Err(KernelError::new(
            "GEOMETRY_PANE_INVALID",
            "Frozen pane split counts must be integers",
        ));
    }
    Ok(())
}

fn freeze_split_count(value: f64, count: u32, axis: &'static str) -> KernelResult<u32> {
    if value.fract() != 0.0 || value < 0.0 || !value.is_finite() {
        return Err(KernelError::new(
            "GEOMETRY_PANE_INVALID",
            "Frozen pane split counts must be finite non-negative integers",
        )
        .at(axis));
    }
    Ok(value.min(count as f64) as u32)
}

fn overlap(a: Option<CellRange>, b: Option<CellRange>) -> bool {
    match (a, b) {
        (Some(a), Some(b)) => {
            a.start_row <= b.end_row
                && b.start_row <= a.end_row
                && a.start_column <= b.end_column
                && b.start_column <= a.end_column
        }
        _ => false,
    }
}
fn clamp_range(
    id: PaneId,
    r: Option<CellRange>,
    fy: u32,
    fx: u32,
    rows: u32,
    cols: u32,
) -> Option<CellRange> {
    let r = r?;
    let (rs, re) = match id {
        PaneId::TopLeft | PaneId::TopRight => (0, fy.saturating_sub(1)),
        _ => (fy, rows.saturating_sub(1)),
    };
    let (cs, ce) = match id {
        PaneId::TopLeft | PaneId::BottomLeft => (0, fx.saturating_sub(1)),
        _ => (fx, cols.saturating_sub(1)),
    };
    let n = CellRange {
        start_row: r.start_row.max(rs),
        end_row: r.end_row.min(re),
        start_column: r.start_column.max(cs),
        end_column: r.end_column.min(ce),
    };
    (n.start_row <= n.end_row && n.start_column <= n.end_column).then_some(n)
}

fn owning_pane<'a>(
    map: &'a PaneMap,
    axis: HeaderAxis,
    index: u32,
    coordinate: f64,
) -> Option<&'a RenderPane> {
    let in_range = |pane: &&RenderPane| {
        pane.visible_range
            .map(|range| match axis {
                HeaderAxis::Row => index >= range.start_row && index <= range.end_row,
                HeaderAxis::Column => index >= range.start_column && index <= range.end_column,
                HeaderAxis::Corner => false,
            })
            .unwrap_or(false)
    };
    map.panes
        .iter()
        .find(in_range)
        .or_else(|| {
            map.panes.iter().find(|pane| {
                let (origin, extent) = match axis {
                    HeaderAxis::Row => (pane.content_origin.y, pane.screen_rect.height),
                    HeaderAxis::Column => (pane.content_origin.x, pane.screen_rect.width),
                    HeaderAxis::Corner => (0.0, 0.0),
                };
                coordinate >= origin && coordinate < origin + extent
            })
        })
        .or_else(|| map.panes.first())
}

pub fn compute_pane_map(request: &GeometryRequest) -> KernelResult<GeometryResponse> {
    let g = Geometry::from(request)?;
    Ok(GeometryResponse {
        pane_map: g.pane_map(request.pane.as_ref())?,
    })
}
pub fn hit_test(request: &GeometryRequest, point: Point) -> KernelResult<HitTestResponse> {
    let g = Geometry::from(request)?;
    if !point.x.is_finite() || !point.y.is_finite() {
        return Err(KernelError::new(
            "GEOMETRY_POINT_INVALID",
            "Hit-test point is invalid",
        ));
    }
    let map = g.pane_map(request.pane.as_ref())?;
    let pane = map.panes.iter().find(|p| {
        point.x >= p.screen_rect.x
            && point.x < p.screen_rect.x + p.screen_rect.width
            && point.y >= p.screen_rect.y
            && point.y < p.screen_rect.y + p.screen_rect.height
    });
    let Some(p) = pane else {
        return Ok(HitTestResponse {
            address: None,
            pane: None,
        });
    };
    let address = g
        .columns
        .index_at(p.content_origin.x + point.x - p.screen_rect.x)
        .and_then(|column| {
            g.rows
                .index_at(p.content_origin.y + point.y - p.screen_rect.y)
                .map(|row| CellAddress {
                    sheet_id: g.sheet_id.clone(),
                    row,
                    column,
                })
        });
    Ok(HitTestResponse {
        address,
        pane: Some(p.id),
    })
}
pub fn cell_rect(request: &GeometryRequest, address: &CellAddress) -> KernelResult<Rect> {
    let g = Geometry::from(request)?;
    address.validate()?;
    if address.sheet_id != g.sheet_id {
        return Err(KernelError::new(
            "GEOMETRY_SHEET_MISMATCH",
            "Cell belongs to another sheet",
        ));
    }
    if g.rows.hidden.contains(&address.row) || g.columns.hidden.contains(&address.column) {
        return Err(KernelError::new(
            "GEOMETRY_CELL_HIDDEN",
            "Hidden cells have no screen rectangle",
        ));
    }
    let map = g.pane_map(request.pane.as_ref())?;
    let p = map
        .panes
        .iter()
        .find(|p| {
            p.visible_range
                .map(|r| {
                    address.row >= r.start_row
                        && address.row <= r.end_row
                        && address.column >= r.start_column
                        && address.column <= r.end_column
                })
                .unwrap_or(false)
        })
        .ok_or_else(|| {
            KernelError::new("GEOMETRY_CELL_NOT_VISIBLE", "Cell is outside the viewport")
        })?;
    Ok(Rect {
        x: p.screen_rect.x + g.columns.total_before(address.column) - p.content_origin.x,
        y: p.screen_rect.y + g.rows.total_before(address.row) - p.content_origin.y,
        width: g.columns.size(address.column),
        height: g.rows.size(address.row),
    })
}
pub fn header_rect(
    request: &GeometryRequest,
    axis: HeaderAxis,
    index: Option<u32>,
) -> KernelResult<HeaderRectResponse> {
    let g = Geometry::from(request)?;
    let x = g.origin.x;
    let y = g.origin.y;
    match axis {
        HeaderAxis::Corner => Ok(HeaderRectResponse {
            rect: Rect {
                x: 0.0,
                y: 0.0,
                width: x,
                height: y,
            },
            axis,
            index,
        }),
        HeaderAxis::Row => {
            let i = index.ok_or_else(|| {
                KernelError::new(
                    "GEOMETRY_HEADER_INDEX_REQUIRED",
                    "Row header requires an index",
                )
            })?;
            if i >= g.rows.count {
                return Err(KernelError::new(
                    "GEOMETRY_HEADER_INDEX_INVALID",
                    "Row header index is outside the worksheet",
                ));
            }
            let coordinate = g.rows.total_before(i);
            let map = g.pane_map(request.pane.as_ref())?;
            let pane = owning_pane(&map, axis, i, coordinate);
            let (screen_y, content_y) = pane
                .map(|pane| (pane.screen_rect.y, pane.content_origin.y))
                .unwrap_or((y, g.viewport.scroll_y));
            Ok(HeaderRectResponse {
                rect: Rect {
                    x: 0.0,
                    y: screen_y + coordinate - content_y,
                    width: x,
                    height: g.rows.size(i),
                },
                axis,
                index: Some(i),
            })
        }
        HeaderAxis::Column => {
            let i = index.ok_or_else(|| {
                KernelError::new(
                    "GEOMETRY_HEADER_INDEX_REQUIRED",
                    "Column header requires an index",
                )
            })?;
            if i >= g.columns.count {
                return Err(KernelError::new(
                    "GEOMETRY_HEADER_INDEX_INVALID",
                    "Column header index is outside the worksheet",
                ));
            }
            let coordinate = g.columns.total_before(i);
            let map = g.pane_map(request.pane.as_ref())?;
            let pane = owning_pane(&map, axis, i, coordinate);
            let (screen_x, content_x) = pane
                .map(|pane| (pane.screen_rect.x, pane.content_origin.x))
                .unwrap_or((x, g.viewport.scroll_x));
            Ok(HeaderRectResponse {
                rect: Rect {
                    x: screen_x + coordinate - content_x,
                    y: 0.0,
                    width: g.columns.size(i),
                    height: y,
                },
                axis,
                index: Some(i),
            })
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn req() -> GeometryRequest {
        GeometryRequest {
            sheet_id: "s".into(),
            row_count: 100,
            column_count: 20,
            default_row_height_px: 20.0,
            default_column_width_px: 50.0,
            row_heights_px: BTreeMap::new(),
            column_widths_px: BTreeMap::new(),
            hidden_rows: BTreeSet::new(),
            hidden_columns: BTreeSet::new(),
            zoom: 1.0,
            viewport: Viewport {
                width: 500.0,
                height: 300.0,
                scroll_x: 0.0,
                scroll_y: 0.0,
                device_pixel_ratio: 1.0,
            },
            pane: None,
            header_offset: None,
        }
    }
    #[test]
    fn sparse_hidden_and_zoom() {
        let mut r = req();
        r.row_count = 1_000_000;
        r.hidden_rows.insert(2);
        r.row_heights_px.insert(10, 40.0);
        r.zoom = 1.25;
        let a = CellAddress {
            sheet_id: "s".into(),
            row: 10,
            column: 0,
        };
        let rect = cell_rect(&r, &a).unwrap();
        assert_eq!(rect.height, 50.0);
        assert_eq!(rect.y, 245.0);
    }
    #[test]
    fn frozen_ranges_disjoint_and_hit_matches_cell() {
        let mut r = req();
        r.pane = Some(PaneLayout {
            kind: PaneKind::Frozen,
            x_split: 2.0,
            y_split: 2.0,
            start_row: 2,
            start_column: 2,
            active_pane: None,
            state: PaneState::Frozen,
        });
        let map = compute_pane_map(&r).unwrap().pane_map;
        for a in 0..map.panes.len() {
            for b in a + 1..map.panes.len() {
                assert!(!overlap(
                    map.panes[a].visible_range,
                    map.panes[b].visible_range
                ));
            }
        }
        let p = Point { x: 45.0, y: 25.0 };
        let hit = hit_test(&r, p).unwrap().address.unwrap();
        let rect = cell_rect(&r, &hit).unwrap();
        assert!(rect.x <= p.x && p.x < rect.x + rect.width);
    }

    #[test]
    fn single_axis_freeze_does_not_reserve_the_unfrozen_axis() {
        let mut r = req();
        r.pane = Some(PaneLayout {
            kind: PaneKind::Frozen,
            x_split: 2.0,
            y_split: 0.0,
            start_row: 0,
            start_column: 2,
            active_pane: None,
            state: PaneState::Frozen,
        });
        let map = compute_pane_map(&r).unwrap().pane_map;
        assert!(map
            .panes
            .iter()
            .all(|pane| { !matches!(pane.id, PaneId::TopLeft | PaneId::TopRight) }));
        assert_eq!(
            map.panes.iter().map(|pane| pane.id).collect::<Vec<_>>(),
            vec![PaneId::BottomLeft, PaneId::Main]
        );
        assert!(map
            .panes
            .iter()
            .all(|pane| pane.screen_rect.height == 280.0));

        r.pane = Some(PaneLayout {
            kind: PaneKind::Frozen,
            x_split: 0.0,
            y_split: 2.0,
            start_row: 2,
            start_column: 0,
            active_pane: None,
            state: PaneState::Frozen,
        });
        let map = compute_pane_map(&r).unwrap().pane_map;
        assert!(map
            .panes
            .iter()
            .all(|pane| { !matches!(pane.id, PaneId::TopLeft | PaneId::BottomLeft) }));
        assert_eq!(
            map.panes.iter().map(|pane| pane.id).collect::<Vec<_>>(),
            vec![PaneId::TopRight, PaneId::Main]
        );
        assert!(map.panes.iter().all(|pane| pane.screen_rect.width == 461.0));
    }

    #[test]
    fn split_positions_are_converted_from_twentieths_of_a_point_to_css_pixels() {
        let mut r = req();
        r.pane = Some(PaneLayout {
            kind: PaneKind::Split,
            x_split: 1440.5,
            y_split: 720.25,
            start_row: 0,
            start_column: 0,
            active_pane: None,
            state: PaneState::Split,
        });
        let main = compute_pane_map(&r)
            .unwrap()
            .pane_map
            .panes
            .into_iter()
            .find(|pane| pane.id == PaneId::Main)
            .unwrap();
        assert!((main.content_origin.x - 96.03333333333333).abs() < 1e-12);
        assert!((main.content_origin.y - 48.016666666666666).abs() < 1e-12);
    }

    #[test]
    fn split_layout_contract_accepts_fractional_twentieths() {
        let pane: PaneLayout = serde_json::from_value(serde_json::json!({
            "kind": "split",
            "xSplit": 15.5,
            "ySplit": 30.25,
            "startRow": 0,
            "startColumn": 0,
            "state": "split"
        }))
        .unwrap();
        assert_eq!(pane.x_split, 15.5);
        assert_eq!(pane.y_split, 30.25);
    }

    #[test]
    fn frozen_split_counts_reject_fractional_positions() {
        let mut r = req();
        r.pane = Some(PaneLayout {
            kind: PaneKind::Frozen,
            x_split: 1.5,
            y_split: 0.0,
            start_row: 0,
            start_column: 1,
            active_pane: None,
            state: PaneState::Frozen,
        });
        assert_eq!(
            compute_pane_map(&r).unwrap_err().code,
            "GEOMETRY_PANE_INVALID"
        );
    }

    #[test]
    fn saved_frozen_start_coordinates_are_not_offset_twice() {
        let mut r = req();
        r.viewport.scroll_x = 250.0;
        r.viewport.scroll_y = 80.0;
        r.pane = Some(PaneLayout {
            kind: PaneKind::Frozen,
            x_split: 2.0,
            y_split: 2.0,
            start_row: 4,
            start_column: 5,
            active_pane: None,
            state: PaneState::Frozen,
        });
        let main = compute_pane_map(&r)
            .unwrap()
            .pane_map
            .panes
            .into_iter()
            .find(|pane| pane.id == PaneId::Main)
            .unwrap();
        assert_eq!(main.content_origin, Point { x: 250.0, y: 80.0 });
    }

    #[test]
    fn frozen_header_rect_uses_the_owning_pane_origin() {
        let mut r = req();
        r.viewport.scroll_x = 100.0;
        r.viewport.scroll_y = 40.0;
        r.pane = Some(PaneLayout {
            kind: PaneKind::Frozen,
            x_split: 2.0,
            y_split: 2.0,
            start_row: 2,
            start_column: 2,
            active_pane: None,
            state: PaneState::Frozen,
        });

        let frozen_row = header_rect(&r, HeaderAxis::Row, Some(0)).unwrap();
        let scrolling_row = header_rect(&r, HeaderAxis::Row, Some(2)).unwrap();
        assert_eq!(frozen_row.rect.y, 20.0);
        assert_eq!(scrolling_row.rect.y, 60.0);

        let frozen_column = header_rect(&r, HeaderAxis::Column, Some(0)).unwrap();
        let scrolling_column = header_rect(&r, HeaderAxis::Column, Some(2)).unwrap();
        assert_eq!(frozen_column.rect.x, 39.0);
        assert_eq!(scrolling_column.rect.x, 139.0);
    }

    #[test]
    fn rejects_invalid_and_hidden() {
        let mut r = req();
        r.hidden_rows.insert(100);
        assert!(compute_pane_map(&r).is_err());
        let mut r = req();
        r.hidden_rows.insert(1);
        let a = CellAddress {
            sheet_id: "s".into(),
            row: 1,
            column: 0,
        };
        assert_eq!(cell_rect(&r, &a).unwrap_err().code, "GEOMETRY_CELL_HIDDEN");
    }
    #[test]
    fn header_origin_keeps_row_one_out_of_column_header() {
        let r = req();
        let h = header_rect(&r, HeaderAxis::Column, Some(0)).unwrap();
        assert_eq!(h.rect.y, 0.0);
        let c = cell_rect(
            &r,
            &CellAddress {
                sheet_id: "s".into(),
                row: 0,
                column: 0,
            },
        )
        .unwrap();
        assert_eq!(c.y, 20.0);
    }

    #[test]
    fn viewport_ending_at_hidden_sheet_extent_keeps_a_visible_range() {
        let mut r = req();
        r.row_count = 6;
        r.column_count = 6;
        r.viewport.width = 220.0;
        r.viewport.height = 120.0;
        r.hidden_rows.extend([1, 2]);
        r.hidden_columns.extend([1, 2]);
        let pane = compute_pane_map(&r).unwrap().pane_map.panes.remove(0);
        assert_eq!(
            pane.visible_range,
            Some(CellRange {
                start_row: 0,
                end_row: 5,
                start_column: 0,
                end_column: 5,
            })
        );
    }
}
