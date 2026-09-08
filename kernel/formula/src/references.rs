//! Structural reference transformations over the canonical source AST.
//! Only changed reference spans are replaced; strings and Unicode source remain intact.
use crate::editor::{self, Expr, Node, ParsedCellReference};
use kernel_core::{KernelError, KernelResult, MAX_COLUMNS, MAX_ROWS, RangeRef};
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Axis {
    Row,
    Column,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Direction {
    Insert,
    Delete,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetIdentity {
    pub id: String,
    pub name: String,
}
#[derive(Clone)]
struct Area {
    start: ParsedCellReference,
    end: ParsedCellReference,
    domain: u8,
    end_explicit: bool,
}
fn bounds() -> KernelError {
    KernelError::new(
        "STRUCTURAL_BOUNDS_INVALID",
        "Formula reference exceeds worksheet bounds",
    )
}
fn area(e: &Expr) -> Option<Area> {
    let base = |sheet_id: Option<String>, row, column, absolute_row, absolute_column| {
        ParsedCellReference {
            sheet_id,
            row,
            column,
            absolute_row,
            absolute_column,
        }
    };
    match &e.node {
        Node::CellReference { reference } => Some(Area {
            start: reference.clone(),
            end: reference.clone(),
            domain: 0,
            end_explicit: false,
        }),
        Node::RangeReference { start, end } => {
            if let (Node::CellReference { reference: a }, Node::CellReference { reference: b }) =
                (&start.node, &end.node)
            {
                let mut b = b.clone();
                let explicit = b.sheet_id.is_some();
                if b.sheet_id.is_none() {
                    b.sheet_id = a.sheet_id.clone();
                }
                Some(Area {
                    start: a.clone(),
                    end: b,
                    domain: 0,
                    end_explicit: explicit,
                })
            } else {
                None
            }
        }
        Node::WholeColumnReference {
            sheet_id,
            start_column,
            end_column,
            absolute_start,
            absolute_end,
        } => Some(Area {
            start: base(sheet_id.clone(), 0, *start_column, true, *absolute_start),
            end: base(
                sheet_id.clone(),
                MAX_ROWS - 1,
                *end_column,
                true,
                *absolute_end,
            ),
            domain: 1,
            end_explicit: false,
        }),
        Node::WholeRowReference {
            sheet_id,
            start_row,
            end_row,
            absolute_start,
            absolute_end,
        } => Some(Area {
            start: base(sheet_id.clone(), *start_row, 0, *absolute_start, true),
            end: base(
                sheet_id.clone(),
                *end_row,
                MAX_COLUMNS - 1,
                *absolute_end,
                true,
            ),
            domain: 2,
            end_explicit: false,
        }),
        _ => None,
    }
}
fn rebuild(e: &Expr, a: Area) -> Expr {
    let mut out = e.clone();
    out.node = match &e.node {
        Node::CellReference { .. } => Node::CellReference { reference: a.start },
        Node::RangeReference { start, end } => {
            let mut s = start.as_ref().clone();
            let mut t = end.as_ref().clone();
            let mut b = a.end;
            if !a.end_explicit && b.sheet_id == a.start.sheet_id {
                b.sheet_id = None;
            }
            s.node = Node::CellReference { reference: a.start };
            t.node = Node::CellReference { reference: b };
            Node::RangeReference {
                start: Box::new(s),
                end: Box::new(t),
            }
        }
        Node::WholeColumnReference { .. } => Node::WholeColumnReference {
            sheet_id: a.start.sheet_id,
            start_column: a.start.column,
            end_column: a.end.column,
            absolute_start: a.start.absolute_column,
            absolute_end: a.end.absolute_column,
        },
        Node::WholeRowReference { .. } => Node::WholeRowReference {
            sheet_id: a.start.sheet_id,
            start_row: a.start.row,
            end_row: a.end.row,
            absolute_start: a.start.absolute_row,
            absolute_end: a.end.absolute_row,
        },
        _ => unreachable!(),
    };
    out
}
fn rewrite(
    formula: &str,
    mapper: &mut impl FnMut(&Expr, Area) -> KernelResult<Expr>,
) -> KernelResult<String> {
    if !formula.trim_start().starts_with('=') {
        return Ok(formula.into());
    }
    let ast = editor::parse_source(formula)?;
    let mut edits = Vec::new();
    fn visit(
        e: &Expr,
        mapper: &mut impl FnMut(&Expr, Area) -> KernelResult<Expr>,
        edits: &mut Vec<(u32, u32, String)>,
    ) -> KernelResult<()> {
        if let Some(a) = area(e) {
            let mapped = mapper(e, a)?;
            if mapped != *e {
                edits.push((e.span.start, e.span.end, editor::format_editor(&mapped)));
            }
            return Ok(());
        }
        for c in editor::children(e) {
            visit(c, mapper, edits)?;
        }
        Ok(())
    }
    visit(&ast, mapper, &mut edits)?;
    // UTF-16→byte boundary map is built once, never by indexing UTF-8 bytes as chars.
    let mut boundaries = std::collections::BTreeMap::new();
    let mut units = 0u32;
    for (byte, c) in formula.char_indices() {
        boundaries.insert(units, byte);
        units += c.len_utf16() as u32;
    }
    boundaries.insert(units, formula.len());
    let mut out = formula.to_owned();
    for (start, end, text) in edits.into_iter().rev() {
        let a = *boundaries.get(&start).ok_or_else(|| {
            KernelError::new(
                "FORMULA_SPAN_INVALID",
                "Reference start is not a Unicode boundary",
            )
        })?;
        let b = *boundaries.get(&end).ok_or_else(|| {
            KernelError::new(
                "FORMULA_SPAN_INVALID",
                "Reference end is not a Unicode boundary",
            )
        })?;
        out.replace_range(a..b, &text);
    }
    Ok(out)
}
fn same(a: &str, b: &str) -> bool {
    a.eq_ignore_ascii_case(b)
}
fn belongs(r: &ParsedCellReference, owner: &SheetIdentity, target: &SheetIdentity) -> bool {
    match &r.sheet_id {
        None => owner.id == target.id,
        Some(s) => same(s, &target.id) || same(s, &target.name),
    }
}
fn coordinate(r: &ParsedCellReference, axis: Axis) -> u32 {
    if axis == Axis::Row { r.row } else { r.column }
}
fn set_coordinate(r: &mut ParsedCellReference, axis: Axis, p: u32) {
    if axis == Axis::Row {
        r.row = p
    } else {
        r.column = p
    }
}
fn limit(axis: Axis) -> u32 {
    if axis == Axis::Row {
        MAX_ROWS
    } else {
        MAX_COLUMNS
    }
}
fn interval(
    a: u32,
    b: u32,
    at: u32,
    count: u32,
    direction: Direction,
    limit: u32,
) -> KernelResult<Option<(u32, u32)>> {
    let reverse = a > b;
    let lo = a.min(b);
    let hi = a.max(b);
    let result = match direction {
        Direction::Insert => {
            let map = |v: u32| {
                if v < at {
                    Ok(v)
                } else {
                    v.checked_add(count)
                        .filter(|n| *n < limit)
                        .ok_or_else(bounds)
                }
            };
            Some((map(lo)?, map(hi)?))
        }
        Direction::Delete => {
            let deleted_end = at.checked_add(count - 1).ok_or_else(bounds)?;
            if hi < at {
                Some((lo, hi))
            } else if lo > deleted_end {
                Some((lo - count, hi - count))
            } else if lo >= at && hi <= deleted_end {
                None
            } else {
                let low = if lo < at { lo } else { at };
                let high = if hi > deleted_end { hi - count } else { at - 1 };
                Some((low, high))
            }
        }
    };
    Ok(result.map(|(a, b)| if reverse { (b, a) } else { (a, b) }))
}
fn overlaps(a: &Area, r: &RangeRef) -> bool {
    a.start.row.min(a.end.row) <= r.end_row
        && a.start.row.max(a.end.row) >= r.start_row
        && a.start.column.min(a.end.column) <= r.end_column
        && a.start.column.max(a.end.column) >= r.start_column
}
fn inside(a: &Area, r: &RangeRef) -> bool {
    a.start.row.min(a.end.row) >= r.start_row
        && a.start.row.max(a.end.row) <= r.end_row
        && a.start.column.min(a.end.column) >= r.start_column
        && a.start.column.max(a.end.column) <= r.end_column
}
pub fn remap_axis(
    formula: &str,
    owner: &SheetIdentity,
    target: &SheetIdentity,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
) -> KernelResult<String> {
    remap_axis_in_scope(formula, owner, target, axis, at, count, direction, None)
}
pub fn remap_axis_in_scope(
    formula: &str,
    owner: &SheetIdentity,
    target: &SheetIdentity,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
    scope: Option<&RangeRef>,
) -> KernelResult<String> {
    let max = limit(axis);
    if count == 0 || at >= max || count > max - at {
        return Err(bounds());
    }
    if let Some(s) = scope {
        s.validate()?;
        if s.sheet_id != target.id {
            return Err(KernelError::new(
                "STRUCTURAL_SCOPE_INVALID",
                "Formula transform scope must belong to target worksheet",
            ));
        }
    }
    rewrite(formula, &mut |e, mut a| {
        if !belongs(&a.start, owner, target)
            || a.domain == 1 && axis == Axis::Row
            || a.domain == 2 && axis == Axis::Column
        {
            return Ok(e.clone());
        }
        if let Some(scope) = scope {
            if !overlaps(&a, scope) {
                return Ok(e.clone());
            }
            if !inside(&a, scope) {
                return Err(KernelError::new(
                    "UNSUPPORTED_FEATURE",
                    "Scoped shift would split a formula range; transaction rejected",
                ));
            }
        }
        let Some((start, end)) = interval(
            coordinate(&a.start, axis),
            coordinate(&a.end, axis),
            at,
            count,
            direction,
            max,
        )?
        else {
            return Ok(editor::invalid(e));
        };
        set_coordinate(&mut a.start, axis, start);
        set_coordinate(&mut a.end, axis, end);
        Ok(rebuild(e, a))
    })
}
pub fn offset(formula: &str, row_offset: i32, column_offset: i32) -> KernelResult<String> {
    rewrite(formula, &mut |e, _| {
        Ok(editor::offset_editor(e, row_offset, column_offset))
    })
}
pub fn remap_moved_region(
    formula: &str,
    owner: &SheetIdentity,
    target: &SheetIdentity,
    selection: &RangeRef,
    row_delta: i32,
    column_delta: i32,
) -> KernelResult<String> {
    selection.validate()?;
    if selection.sheet_id != target.id {
        return Err(KernelError::new(
            "STRUCTURAL_SCOPE_INVALID",
            "Moved selection must belong to target worksheet",
        ));
    }
    rewrite(formula, &mut |e, mut a| {
        if !belongs(&a.start, owner, target) || !overlaps(&a, selection) {
            return Ok(e.clone());
        }
        if !inside(&a, selection) {
            return Err(KernelError::new(
                "UNSUPPORTED_FEATURE",
                "Moving part of a referenced range would create a non-contiguous reference",
            ));
        }
        for r in [&mut a.start, &mut a.end] {
            let row = r.row as i64 + row_delta as i64;
            let column = r.column as i64 + column_delta as i64;
            if row < 0 || row >= MAX_ROWS as i64 || column < 0 || column >= MAX_COLUMNS as i64 {
                return Err(bounds());
            }
            r.row = row as u32;
            r.column = column as u32;
        }
        Ok(rebuild(e, a))
    })
}
pub fn rename_sheet(formula: &str, old_name: &str, new_name: &str) -> KernelResult<String> {
    if old_name.trim().is_empty() || new_name.trim().is_empty() {
        return Err(KernelError::new(
            "SHEET_IDENTITY_INVALID",
            "Worksheet names are required",
        ));
    }
    rewrite(formula, &mut |e, mut a| {
        for r in [&mut a.start, &mut a.end] {
            if r.sheet_id.as_ref().is_some_and(|s| same(s, old_name)) {
                r.sheet_id = Some(new_name.into());
            }
        }
        Ok(rebuild(e, a))
    })
}
pub fn invalidate_sheet(formula: &str, sheet_id: &str, sheet_name: &str) -> KernelResult<String> {
    if sheet_id.trim().is_empty() || sheet_name.trim().is_empty() {
        return Err(KernelError::new(
            "SHEET_IDENTITY_INVALID",
            "Worksheet identity is required",
        ));
    }
    rewrite(formula, &mut |e, a| {
        if [&a.start, &a.end].iter().any(|r| {
            r.sheet_id
                .as_ref()
                .is_some_and(|s| same(s, sheet_id) || same(s, sheet_name))
        }) {
            Ok(editor::invalid(e))
        } else {
            Ok(e.clone())
        }
    })
}
#[cfg(test)]
mod tests {
    use super::*;
    fn sheet() -> SheetIdentity {
        SheetIdentity {
            id: "sheet-1".into(),
            name: "Sheet1".into(),
        }
    }
    #[test]
    fn range_deletion_shrinks_and_preserves_direction() {
        let s = sheet();
        for (formula, at, count, want) in [
            ("=A1:A10", 0, 1, "=A1:A9"),
            ("=A1:A10", 9, 1, "=A1:A9"),
            ("=A1:A10", 4, 2, "=A1:A8"),
            ("=A10:A1", 4, 2, "=A8:A1"),
        ] {
            assert_eq!(
                remap_axis(formula, &s, &s, Axis::Row, at, count, Direction::Delete).unwrap(),
                want
            );
        }
    }
    #[test]
    fn deleted_cells_become_ref() {
        let s = sheet();
        assert_eq!(
            remap_axis("=A1:A2", &s, &s, Axis::Row, 0, 2, Direction::Delete).unwrap(),
            "=#REF!"
        );
        assert_eq!(offset("=A1", -1, 0).unwrap(), "=#REF!");
    }
    #[test]
    fn unicode_strings_names_and_qualified_ranges_are_preserved() {
        assert_eq!(
            rename_sheet(
                "=SUM('数据表😀'!$A$1:B2)+\"数据表😀 A1\"",
                "数据表😀",
                "新表"
            )
            .unwrap(),
            "=SUM(新表!$A$1:B2)+\"数据表😀 A1\""
        );
        assert_eq!(
            offset("=\"😀A1\"&名称+A1", 1, 1).unwrap(),
            "=\"😀A1\"&名称+B2"
        );
    }
    #[test]
    fn whole_columns_and_rows_transform_as_domains() {
        let s = sheet();
        assert_eq!(
            remap_axis("=$A:$C", &s, &s, Axis::Column, 1, 1, Direction::Insert).unwrap(),
            "=$A:$D"
        );
        assert_eq!(
            remap_axis("=1:5", &s, &s, Axis::Row, 2, 1, Direction::Delete).unwrap(),
            "=1:4"
        );
        assert_eq!(
            remap_axis("=A:C", &s, &s, Axis::Row, 2, 1, Direction::Insert).unwrap(),
            "=A:C"
        );
    }
    #[test]
    fn scoped_range_split_and_invalid_count_reject() {
        let s = sheet();
        assert!(remap_axis("=A1", &s, &s, Axis::Row, 0, 0, Direction::Insert).is_err());
        let scope = RangeRef {
            sheet_id: s.id.clone(),
            start_row: 0,
            end_row: 1,
            start_column: 0,
            end_column: 0,
        };
        assert_eq!(
            remap_axis_in_scope(
                "=A2+A10",
                &s,
                &s,
                Axis::Row,
                1,
                1,
                Direction::Insert,
                Some(&scope)
            )
            .unwrap(),
            "=A3+A10"
        );
        assert_eq!(
            remap_axis_in_scope(
                "=A1:B2",
                &s,
                &s,
                Axis::Row,
                1,
                1,
                Direction::Insert,
                Some(&scope)
            )
            .unwrap_err()
            .code,
            "UNSUPPORTED_FEATURE"
        );
    }
    #[test]
    fn move_updates_absolute_and_cross_sheet_observers() {
        let s = sheet();
        let observer = SheetIdentity {
            id: "other".into(),
            name: "Other".into(),
        };
        let selection = RangeRef {
            sheet_id: s.id.clone(),
            start_row: 0,
            end_row: 1,
            start_column: 0,
            end_column: 1,
        };
        assert_eq!(
            remap_moved_region("=Sheet1!$A$1:B2+A1", &observer, &s, &selection, 2, 3).unwrap(),
            "=Sheet1!$D$3:E4+A1"
        );
    }
}
