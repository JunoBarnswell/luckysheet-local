//! Workbook and worksheet structural mutation reducers.
//!
//! The reducer owns the complete structural transaction: authored cells,
//! formulas, rules and ranges, object anchors, worksheet dimensions and
//! worksheet lifecycle.  Metadata is deliberately transformed in the same
//! pass as cells; there is no UI-only repair path.

use crate::Transaction;
use kernel_core::{
    Cell, CellAddress, KernelError, KernelResult, MAX_COLUMNS, MAX_ROWS, RangeRef, SheetManifest,
};
use kernel_formula::references::{self, Axis, Direction, SheetIdentity};
use serde_json::{Map, Value};
use std::collections::{BTreeMap, BTreeSet, HashMap};

const STRUCTURAL_IDS: &[&str] = &[
    "rows.inserted",
    "rows.deleted",
    "columns.inserted",
    "columns.deleted",
    "cells.inserted",
    "cells.deleted",
    "cells.inserted.restore",
    "cells.deleted.restore",
    "rows.permuted",
];
const WORKBOOK_IDS: &[&str] = &[
    "sheet.add",
    "sheet.remove",
    "sheet.rename",
    "sheet.duplicated",
    "sheet.restore",
    "hyperlink.set",
    "hyperlink.remove",
    "sheet.reordered",
];

pub(crate) fn apply(
    tx: &mut Transaction,
    id: &str,
    sheet_id: &str,
    p: &Value,
) -> KernelResult<bool> {
    if STRUCTURAL_IDS.contains(&id) {
        match id {
            "rows.inserted" => axis(tx, sheet_id, p, Axis::Row, Direction::Insert),
            "rows.deleted" => axis(tx, sheet_id, p, Axis::Row, Direction::Delete),
            "columns.inserted" => axis(tx, sheet_id, p, Axis::Column, Direction::Insert),
            "columns.deleted" => axis(tx, sheet_id, p, Axis::Column, Direction::Delete),
            "cells.inserted" | "cells.deleted" => cell_shift(tx, sheet_id, p),
            "cells.inserted.restore" | "cells.deleted.restore" => restore_shift(tx, sheet_id, p),
            "rows.permuted" => permute_rows(tx, sheet_id, p),
            _ => unreachable!(),
        }?;
        return Ok(true);
    }
    if WORKBOOK_IDS.contains(&id) {
        match id {
            "sheet.add" => add_sheet(tx, sheet_id, p),
            "sheet.remove" => remove_sheet(tx, p),
            "sheet.rename" => rename_sheet(tx, sheet_id, p),
            "sheet.duplicated" => duplicate_sheet(tx, p),
            "sheet.restore" => restore_sheet(tx, p),
            "sheet.reordered" => reorder_sheet(tx, sheet_id, p),
            "hyperlink.set" => set_hyperlink(tx, sheet_id, p),
            "hyperlink.remove" => remove_hyperlink(tx, sheet_id, p),
            _ => unreachable!(),
        }?;
        return Ok(true);
    }
    Ok(false)
}

fn axis(
    tx: &mut Transaction,
    sheet_id: &str,
    p: &Value,
    axis: Axis,
    direction: Direction,
) -> KernelResult<()> {
    let at = uint(p, "at")?;
    let count = uint(p, "count")?;
    let (limit, maximum) = {
        let s = tx.sheet(sheet_id)?;
        (
            if axis == Axis::Row {
                s.row_count
            } else {
                s.column_count
            },
            if axis == Axis::Row {
                MAX_ROWS
            } else {
                MAX_COLUMNS
            },
        )
    };
    if count == 0
        || at > limit
        || (direction == Direction::Delete && at.saturating_add(count) > limit)
        || (direction == Direction::Insert && limit.checked_add(count).is_none_or(|v| v > maximum))
    {
        return Err(KernelError::new(
            "STRUCTURAL_BOUNDS_INVALID",
            "Structural axis bounds are invalid",
        )
        .at(sheet_id));
    }
    if direction == Direction::Delete {
        validate_delete_metadata(tx, sheet_id, axis, at, count)?;
    }
    let old = tx.cells(sheet_id)?;
    let mut next = Vec::with_capacity(old.len());
    for (address, cell) in old {
        let value = if axis == Axis::Row {
            address.row
        } else {
            address.column
        };
        let mapped = shift_index(value, at, count, direction);
        if let Some(value) = mapped {
            let mut a = address;
            if axis == Axis::Row {
                a.row = value;
            } else {
                a.column = value;
            }
            next.push((a, cell));
        }
    }
    let identity = sheet_identity(tx, sheet_id)?;
    let final_limit = if direction == Direction::Insert {
        limit + count
    } else {
        (limit - count).max(1)
    };
    if direction == Direction::Insert {
        set_dimension(tx, sheet_id, axis, final_limit)?;
    }
    let current_rows = tx.sheet(sheet_id)?.row_count;
    let current_columns = tx.sheet(sheet_id)?.column_count;
    let bounds = Some((
        if axis == Axis::Row {
            final_limit
        } else {
            current_rows
        },
        if axis == Axis::Column {
            final_limit
        } else {
            current_columns
        },
    ));
    replace_cells_bounded(tx, sheet_id, next, bounds)?;
    if direction == Direction::Delete {
        set_dimension(tx, sheet_id, axis, (limit - count).max(1))?;
    }
    {
        let s = tx.sheet_mut(sheet_id)?;
        shift_metadata(
            &mut s.metadata,
            identity,
            sheet_id,
            axis,
            at,
            count,
            direction,
        )?;
    }
    transform_all_formula_cells(tx, sheet_id, axis, at, count, direction)?;
    transform_manifest_metadata(tx, sheet_id, axis, at, count, direction)?;
    let affected = whole_sheet(tx, sheet_id)?;
    tx.affected.push(affected);
    Ok(())
}

fn cell_shift(tx: &mut Transaction, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let range = own_range(
        tx,
        p.get("range")
            .ok_or_else(|| invalid("Cell shift range is required"))?,
        sheet_id,
    )?;
    let band = own_range(
        tx,
        p.get("affectedBand")
            .ok_or_else(|| invalid("Cell shift affected band is required"))?,
        sheet_id,
    )?;
    let operation = text(p, "operation")?;
    let axis = match text(p, "axis")?.as_str() {
        "row" => Axis::Row,
        "column" => Axis::Column,
        _ => return Err(invalid("Cell shift axis is invalid")),
    };
    if operation != "insert" && operation != "delete" {
        return Err(invalid("Cell shift operation is invalid"));
    }
    let s = tx.sheet(sheet_id)?;
    let expected = if axis == Axis::Row {
        RangeRef {
            sheet_id: sheet_id.into(),
            start_row: range.start_row,
            end_row: s.row_count - 1,
            start_column: range.start_column,
            end_column: range.end_column,
        }
    } else {
        RangeRef {
            sheet_id: sheet_id.into(),
            start_row: range.start_row,
            end_row: range.end_row,
            start_column: range.start_column,
            end_column: s.column_count - 1,
        }
    };
    if expected != band {
        return Err(invalid("Cell shift affected band is not canonical"));
    }
    let count = if axis == Axis::Row {
        range.end_row - range.start_row + 1
    } else {
        range.end_column - range.start_column + 1
    };
    let axis_limit = if axis == Axis::Row {
        s.row_count
    } else {
        s.column_count
    };
    let direction = if operation == "insert" {
        Direction::Insert
    } else {
        Direction::Delete
    };
    if direction == Direction::Delete {
        validate_delete_metadata(
            tx,
            sheet_id,
            axis,
            if axis == Axis::Row {
                range.start_row
            } else {
                range.start_column
            },
            count,
        )?;
    }
    let old = tx.cells(sheet_id)?;
    let mut next = Vec::new();
    for (mut a, c) in old {
        if !band.contains(&a) {
            next.push((a, c));
            continue;
        }
        let v = if axis == Axis::Row { a.row } else { a.column };
        let start = if axis == Axis::Row {
            range.start_row
        } else {
            range.start_column
        };
        let end = if axis == Axis::Row {
            range.end_row
        } else {
            range.end_column
        };
        let mapped = if direction == Direction::Delete && v >= start && v <= end {
            None
        } else if v < start {
            Some(v)
        } else if direction == Direction::Insert {
            Some(
                v.checked_add(count)
                    .ok_or_else(|| invalid("Cell shift exceeds worksheet bounds"))?,
            )
        } else {
            Some(
                v.checked_sub(count)
                    .ok_or_else(|| invalid("Cell shift exceeds worksheet bounds"))?,
            )
        };
        if let Some(v) = mapped {
            if v >= axis_limit {
                return Err(invalid(
                    "Cell shift would discard data outside worksheet bounds",
                ));
            }
            if axis == Axis::Row {
                a.row = v;
            } else {
                a.column = v;
            }
            if band.contains(&a) {
                next.push((a, c));
            }
        }
    }
    replace_cells(tx, sheet_id, next)?;
    let identity = sheet_identity(tx, sheet_id)?;
    {
        let s = tx.sheet_mut(sheet_id)?;
        shift_metadata_scoped(
            &mut s.metadata,
            identity,
            sheet_id,
            axis,
            if axis == Axis::Row {
                range.start_row
            } else {
                range.start_column
            },
            count,
            direction,
            Some(&band),
        )?;
    }
    transform_all_formula_cells_scoped(
        tx,
        sheet_id,
        axis,
        if axis == Axis::Row {
            range.start_row
        } else {
            range.start_column
        },
        count,
        direction,
        Some(&band),
    )?;
    transform_manifest_metadata_scoped(
        tx,
        sheet_id,
        axis,
        if axis == Axis::Row {
            range.start_row
        } else {
            range.start_column
        },
        count,
        direction,
        Some(&band),
    )?;
    tx.affected.push(band);
    Ok(())
}

fn restore_shift(tx: &mut Transaction, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let spec = p
        .get("spec")
        .ok_or_else(|| invalid("Structural restore spec is required"))?;
    let operation = text(spec, "operation")?;
    let axis = match text(spec, "axis")?.as_str() {
        "row" => "row",
        "column" => "column",
        _ => return Err(invalid("Structural restore axis is invalid")),
    };
    let range = own_range(
        tx,
        spec.get("range")
            .ok_or_else(|| invalid("Structural restore range is required"))?,
        sheet_id,
    )?;
    let band = own_range(
        tx,
        spec.get("affectedBand")
            .ok_or_else(|| invalid("Structural restore affected band is required"))?,
        sheet_id,
    )?;
    let mut inverse = p.clone();
    inverse["range"] = serde_json::to_value(&range).unwrap_or(Value::Null);
    inverse["affectedBand"] = serde_json::to_value(&band).unwrap_or(Value::Null);
    inverse["operation"] = Value::String(
        if operation == "insert" {
            "delete"
        } else {
            "insert"
        }
        .into(),
    );
    inverse["axis"] = Value::String(axis.into());
    cell_shift(tx, sheet_id, &inverse)?;
    let old = tx.cells(sheet_id)?;
    for (a, _) in old {
        if band.contains(&a) {
            tx.write(sheet_id, a.row, a.column, None)?;
        }
    }
    let values = p
        .get("cells")
        .and_then(Value::as_array)
        .ok_or_else(|| invalid("Structural restore cells must be an array"))?;
    for entry in values {
        let row = uint(entry, "row")?;
        let column = uint(entry, "column")?;
        let a = CellAddress {
            sheet_id: sheet_id.into(),
            row,
            column,
        };
        if !band.contains(&a) {
            return Err(invalid("Structural restore cell is outside affected band"));
        }
        let cell: Cell = serde_json::from_value(
            entry
                .get("cell")
                .cloned()
                .ok_or_else(|| invalid("Structural restore cell payload is required"))?,
        )
        .map_err(|e| invalid(e.to_string()))?;
        tx.write(sheet_id, row, column, Some(cell))?;
    }
    Ok(())
}

fn permute_rows(tx: &mut Transaction, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let range = own_range(
        tx,
        p.get("range")
            .ok_or_else(|| invalid("Row permutation range is required"))?,
        sheet_id,
    )?;
    let rows = p
        .get("sourceRows")
        .and_then(Value::as_array)
        .ok_or_else(|| invalid("Row permutation sourceRows must be an array"))?;
    let expected = (range.end_row - range.start_row + 1) as usize;
    if rows.len() != expected {
        return Err(invalid("Row permutation length does not match range"));
    }
    let mut seen = BTreeSet::new();
    let mut mapping = Vec::with_capacity(rows.len());
    for v in rows {
        let n = v
            .as_u64()
            .ok_or_else(|| invalid("Row permutation value is invalid"))? as u32;
        if n < range.start_row || n > range.end_row || !seen.insert(n) {
            return Err(invalid("Row permutation is not a permutation"));
        }
        mapping.push(n);
    }
    let end_col = p
        .get("affectedColumnEnd")
        .and_then(Value::as_u64)
        .ok_or_else(|| invalid("Rows permutation affected column end is required"))?
        as u32;
    let old = tx.cells(sheet_id)?;
    let mut next = Vec::new();
    for (mut a, c) in old {
        if a.row >= range.start_row && a.row <= range.end_row && a.column <= end_col {
            let source = a.row;
            let target =
                range.start_row + mapping.iter().position(|x| *x == source).unwrap() as u32;
            a.row = target;
        }
        next.push((a, c));
    }
    replace_cells(tx, sheet_id, next)?;
    {
        let s = tx.sheet_mut(sheet_id)?;
        let mut metadata = Value::Object(std::mem::take(&mut s.metadata).into_iter().collect());
        permute_metadata(&mut metadata, sheet_id, &range, &mapping)?;
        s.metadata = metadata.as_object().ok_or_else(|| invalid("Worksheet metadata must remain an object"))?.iter().map(|(k,v)|(k.clone(),v.clone())).collect();
    }
    permute_manifest_metadata(tx, sheet_id, &range, &mapping)?;
    tx.affected.push(RangeRef {
        sheet_id: sheet_id.into(),
        start_row: range.start_row,
        end_row: range.end_row,
        start_column: 0,
        end_column: end_col,
    });
    Ok(())
}

fn add_sheet(tx: &mut Transaction, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let id = text(p, "id")?;
    let name = text(p, "name")?;
    if id != sheet_id || name.trim().is_empty() {
        return Err(invalid("sheet.add identity is invalid"));
    }
    if tx.manifest.sheets.iter().any(|s| s.sheet_id == id) {
        return Err(KernelError::new("CONFLICT", "Sheet already exists").at(id));
    }
    let rows = p
        .get("rowCount")
        .map(|v| dimension(v, true))
        .transpose()?
        .unwrap_or(1000);
    let cols = p
        .get("columnCount")
        .map(|v| dimension(v, false))
        .transpose()?
        .unwrap_or(26);
    tx.manifest.sheets.push(SheetManifest {
        sheet_id: id,
        name,
        row_count: rows,
        column_count: cols,
        metadata: BTreeMap::new(),
    });
    Ok(())
}

fn remove_sheet(tx: &mut Transaction, p: &Value) -> KernelResult<()> {
    let id = text(p, "id")?;
    if tx.manifest.sheets.len() <= 1 {
        return Err(invalid("A workbook must keep at least one worksheet"));
    }
    let index = tx
        .manifest
        .sheets
        .iter()
        .position(|s| s.sheet_id == id)
        .ok_or_else(|| {
            KernelError::new("SHEET_NOT_FOUND", "Worksheet does not exist").at(id.clone())
        })?;
    let name = tx.manifest.sheets[index].name.clone();
    if has_external_reference(tx, &id, &name)? {
        return Err(KernelError::new(
            "CONFLICT",
            "Cannot delete worksheet with external references",
        )
        .at(id));
    }
    tx.manifest.sheets.remove(index);
    tx.manifest.metadata.remove(&format!("sheet:{}", id));
    Ok(())
}

fn rename_sheet(tx: &mut Transaction, mutation_sheet_id: &str, p: &Value) -> KernelResult<()> {
    let id = text(p, "sheetId")?;
    let name = text(p, "name")?;
    if id != mutation_sheet_id || name.trim().is_empty() {
        return Err(invalid("sheet.rename identity is invalid"));
    }
    let old = tx.sheet(&id)?.name.clone();
    if old == name {
        return Ok(());
    }
    for sheet in tx.manifest.sheets.clone() {
        let identity = SheetIdentity {
            id: sheet.sheet_id.clone(),
            name: sheet.name.clone(),
        };
        let target = SheetIdentity {
            id: id.clone(),
            name: old.clone(),
        };
        let cells = tx.cells(&sheet.sheet_id)?;
        for (a, mut c) in cells {
            if let Some(f) = c.formula.take() {
                c.formula = Some(references::rename_sheet(&f, &old, &name)?);
                tx.write(&a.sheet_id, a.row, a.column, Some(c))?;
            }
        }
        let s = tx.sheet_mut(&identity.id)?;
        for value in s.metadata.values_mut() { rewrite_rule_formulas(value, &old, &name); }
    }
    tx.sheet_mut(&id)?.name = name.clone();
    for value in tx.manifest.metadata.values_mut() { rewrite_rule_formulas(value, &old, &name); }
    Ok(())
}

fn duplicate_sheet(tx: &mut Transaction, p: &Value) -> KernelResult<()> {
    let source_id = text(p, "sourceSheetId")?;
    let new_id = text(p, "newId")?;
    let new_name = text(p, "newName")?;
    if tx.manifest.sheets.iter().any(|s| s.sheet_id == new_id) {
        return Err(KernelError::new(
            "CONFLICT",
            "Duplicate sheet identity is invalid",
        ));
    }
    let index = tx
        .manifest
        .sheets
        .iter()
        .position(|s| s.sheet_id == source_id)
        .ok_or_else(|| {
            KernelError::new("SHEET_NOT_FOUND", "Worksheet does not exist").at(source_id.clone())
        })?;
    let source = tx.manifest.sheets[index].clone();
    let source_cells = tx.cells(&source_id)?;
    let mut metadata = source.metadata.clone();
    for value in metadata.values_mut() { remap_sheet_refs(value, &source_id, &new_id, &source.name, &new_name); }
    let copy = SheetManifest {
        sheet_id: new_id.clone(),
        name: new_name.clone(),
        row_count: source.row_count,
        column_count: source.column_count,
        metadata,
    };
    tx.manifest.sheets.insert(index + 1, copy);
    let owner = SheetIdentity {
        id: source_id,
        name: source.name,
    };
    let target = SheetIdentity {
        id: new_id.clone(),
        name: new_name.clone(),
    };
    for (a, mut c) in source_cells {
        let mut na = a;
        na.sheet_id = new_id.clone();
        if let Some(f) = c.formula.take() {
            c.formula = Some(references::rename_sheet(&f, &owner.name, &target.name)?);
        }
        tx.write(&new_id, na.row, na.column, Some(c))?;
    }
    Ok(())
}

fn restore_sheet(tx: &mut Transaction, p: &Value) -> KernelResult<()> {
    let raw = p
        .get("sheet")
        .ok_or_else(|| invalid("sheet.restore requires a canonical sheet snapshot"))?;
    let id = text(raw, "id")?;
    if tx.manifest.sheets.iter().any(|s| s.sheet_id == id) {
        return Err(KernelError::new(
            "CONFLICT",
            "Restored sheet identity is invalid",
        ));
    }
    let name = text(raw, "name")?;
    let rows = dimension(
        raw.get("rowCount")
            .ok_or_else(|| invalid("rowCount is required"))?,
        true,
    )?;
    let cols = dimension(
        raw.get("columnCount")
            .ok_or_else(|| invalid("columnCount is required"))?,
        false,
    )?;
    let metadata = raw
        .get("metadata")
        .and_then(Value::as_object)
        .map(to_btree)
        .unwrap_or_default();
    let index = p
        .get("index")
        .and_then(Value::as_u64)
        .unwrap_or(tx.manifest.sheets.len() as u64)
        .min(tx.manifest.sheets.len() as u64) as usize;
    tx.manifest.sheets.insert(
        index,
        SheetManifest {
            sheet_id: id,
            name,
            row_count: rows,
            column_count: cols,
            metadata,
        },
    );
    Ok(())
}

fn reorder_sheet(tx: &mut Transaction, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let id = p.get("sheetId").and_then(Value::as_str).unwrap_or(sheet_id);
    let from = tx
        .manifest
        .sheets
        .iter()
        .position(|s| s.sheet_id == id)
        .ok_or_else(|| KernelError::new("SHEET_NOT_FOUND", "Worksheet does not exist").at(id))?;
    let to = p
        .get("toIndex")
        .or_else(|| p.get("index"))
        .and_then(Value::as_u64)
        .ok_or_else(|| invalid("sheet.reordered target index is required"))? as usize;
    if to >= tx.manifest.sheets.len() {
        return Err(invalid("sheet.reordered target index is invalid"));
    }
    let s = tx.manifest.sheets.remove(from);
    tx.manifest.sheets.insert(to, s);
    Ok(())
}

fn set_hyperlink(tx: &mut Transaction, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let row = uint(p, "row")?;
    let column = uint(p, "column")?;
    let link = p
        .get("hyperlink")
        .cloned()
        .ok_or_else(|| invalid("hyperlink.set requires a hyperlink"))?;
    if !link.is_object() {
        return Err(invalid("hyperlink must be an object"));
    }
    let s = tx.sheet_mut(sheet_id)?;
    let links = s
        .metadata
        .entry("hyperlinks".into())
        .or_insert_with(|| Value::Array(Vec::new()));
    let array = links
        .as_array_mut()
        .ok_or_else(|| invalid("Worksheet hyperlinks metadata is invalid"))?;
    array.retain(|v| {
        v.get("row").and_then(Value::as_u64) != Some(row as u64)
            || v.get("column").and_then(Value::as_u64) != Some(column as u64)
    });
    let mut entry = Map::new();
    entry.insert("row".into(), Value::from(row));
    entry.insert("column".into(), Value::from(column));
    entry.insert("hyperlink".into(), link);
    array.push(Value::Object(entry));
    tx.affected.push(cell_range(sheet_id, row, column));
    Ok(())
}
fn remove_hyperlink(tx: &mut Transaction, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let row = uint(p, "row")?;
    let column = uint(p, "column")?;
    let s = tx.sheet_mut(sheet_id)?;
    let array = s
        .metadata
        .get_mut("hyperlinks")
        .and_then(Value::as_array_mut)
        .ok_or_else(|| KernelError::new("NOT_FOUND", "Hyperlink does not exist"))?;
    let old = array.len();
    array.retain(|v| {
        v.get("row").and_then(Value::as_u64) != Some(row as u64)
            || v.get("column").and_then(Value::as_u64) != Some(column as u64)
    });
    if old == array.len() {
        return Err(KernelError::new("NOT_FOUND", "Hyperlink does not exist"));
    }
    tx.affected.push(cell_range(sheet_id, row, column));
    Ok(())
}

fn transform_all_formula_cells(
    tx: &mut Transaction,
    target_id: &str,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
) -> KernelResult<()> {
    transform_all_formula_cells_scoped(tx, target_id, axis, at, count, direction, None)
}
fn transform_all_formula_cells_scoped(
    tx: &mut Transaction,
    target_id: &str,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
    scope: Option<&RangeRef>,
) -> KernelResult<()> {
    let sheets = tx.manifest.sheets.clone();
    for owner in sheets {
        let owner_identity = SheetIdentity {
            id: owner.sheet_id.clone(),
            name: owner.name.clone(),
        };
        let target_identity = sheet_identity(tx, target_id)?;
        let cells = tx.cells(&owner.sheet_id)?;
        for (a, mut c) in cells {
            if let Some(f) = c.formula.take() {
                c.formula = Some(references::remap_axis_in_scope(
                    &f,
                    &owner_identity,
                    &target_identity,
                    axis,
                    at,
                    count,
                    direction,
                    scope,
                )?);
                tx.write(&a.sheet_id, a.row, a.column, Some(c))?;
            }
        }
    }
    Ok(())
}

fn transform_manifest_metadata(
    tx: &mut Transaction,
    target_id: &str,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
) -> KernelResult<()> {
    transform_manifest_metadata_scoped(tx, target_id, axis, at, count, direction, None)
}
fn transform_manifest_metadata_scoped(
    tx: &mut Transaction,
    target_id: &str,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
    scope: Option<&RangeRef>,
) -> KernelResult<()> {
    let target = tx.sheet(target_id)?.clone();
    let identity = SheetIdentity {
        id: target.sheet_id,
        name: target.name,
    };
    shift_metadata_scoped(
        &mut tx.manifest.metadata,
        identity,
        target_id,
        axis,
        at,
        count,
        direction,
        scope,
    )
}

fn replace_cells(
    tx: &mut Transaction,
    sheet_id: &str,
    cells: Vec<(CellAddress, Cell)>,
) -> KernelResult<()> {
    replace_cells_bounded(tx, sheet_id, cells, None)
}
fn replace_cells_bounded(
    tx: &mut Transaction,
    sheet_id: &str,
    cells: Vec<(CellAddress, Cell)>,
    bounds: Option<(u32, u32)>,
) -> KernelResult<()> {
    let old = tx.cells(sheet_id)?;
    for (a, _) in old {
        if bounds.is_none_or(|(rows, cols)| a.row < rows && a.column < cols) {
            tx.write(sheet_id, a.row, a.column, None)?;
        }
    }
    for (a, c) in cells {
        tx.write(sheet_id, a.row, a.column, Some(c))?;
    }
    Ok(())
}
fn set_dimension(tx: &mut Transaction, sheet_id: &str, axis: Axis, value: u32) -> KernelResult<()> {
    if value == 0
        || value
            > if axis == Axis::Row {
                MAX_ROWS
            } else {
                MAX_COLUMNS
            }
    {
        return Err(invalid("Structural result exceeds worksheet bounds"));
    }
    let s = tx.sheet_mut(sheet_id)?;
    if axis == Axis::Row {
        s.row_count = value;
    } else {
        s.column_count = value;
    }
    Ok(())
}
fn shift_index(v: u32, at: u32, count: u32, d: Direction) -> Option<u32> {
    match d {
        Direction::Insert => v.checked_add((v >= at).then_some(count).unwrap_or(0)),
        Direction::Delete => {
            if v < at {
                Some(v)
            } else if v < at + count {
                None
            } else {
                Some(v - count)
            }
        }
    }
}
fn whole_sheet(tx: &Transaction, id: &str) -> KernelResult<RangeRef> {
    let s = tx.sheet(id)?;
    Ok(RangeRef {
        sheet_id: id.into(),
        start_row: 0,
        end_row: s.row_count - 1,
        start_column: 0,
        end_column: s.column_count - 1,
    })
}
fn cell_range(id: &str, row: u32, col: u32) -> RangeRef {
    RangeRef {
        sheet_id: id.into(),
        start_row: row,
        end_row: row,
        start_column: col,
        end_column: col,
    }
}
fn own_range(tx: &Transaction, v: &Value, id: &str) -> KernelResult<RangeRef> {
    crate::range(tx, id, v)
}
fn uint(v: &Value, field: &str) -> KernelResult<u32> {
    v.get(field)
        .and_then(Value::as_u64)
        .filter(|x| *x <= u32::MAX as u64)
        .map(|x| x as u32)
        .ok_or_else(|| invalid(format!("{field} is invalid")))
}
fn text(v: &Value, field: &str) -> KernelResult<String> {
    v.get(field)
        .and_then(Value::as_str)
        .filter(|x| !x.trim().is_empty())
        .map(str::to_owned)
        .ok_or_else(|| invalid(format!("{field} is invalid")))
}
fn dimension(v: &Value, row: bool) -> KernelResult<u32> {
    let n = v
        .as_u64()
        .filter(|x| {
            *x > 0
                && *x
                    <= if row {
                        MAX_ROWS as u64
                    } else {
                        MAX_COLUMNS as u64
                    }
        })
        .ok_or_else(|| invalid("Worksheet dimension is invalid"))?;
    Ok(n as u32)
}
fn invalid(v: impl Into<String>) -> KernelError {
    KernelError::new("VALIDATION", v)
}
fn sheet_identity(tx: &Transaction, id: &str) -> KernelResult<SheetIdentity> {
    let s = tx.sheet(id)?;
    Ok(SheetIdentity {
        id: s.sheet_id.clone(),
        name: s.name.clone(),
    })
}
fn to_btree(m: &Map<String, Value>) -> BTreeMap<String, Value> {
    m.iter().map(|(k, v)| (k.clone(), v.clone())).collect()
}

fn shift_metadata(
    metadata: &mut BTreeMap<String, Value>,
    owner: SheetIdentity,
    target_id: &str,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
) -> KernelResult<()> {
    shift_metadata_scoped(metadata, owner, target_id, axis, at, count, direction, None)
}
fn shift_metadata_scoped(
    metadata: &mut BTreeMap<String, Value>,
    owner: SheetIdentity,
    target_id: &str,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
    scope: Option<&RangeRef>,
) -> KernelResult<()> {
    for v in metadata.values_mut() {
        shift_value(
            v,
            Some(&owner),
            target_id,
            axis,
            at,
            count,
            direction,
            scope,
        )?;
    }
    Ok(())
}
fn shift_value(
    v: &mut Value,
    owner: Option<&SheetIdentity>,
    target_id: &str,
    axis: Axis,
    at: u32,
    count: u32,
    direction: Direction,
    scope: Option<&RangeRef>,
) -> KernelResult<()> {
    match v {
        Value::Object(o) => {
            let explicit = o.get("sheetId").and_then(Value::as_str);
            let belongs = explicit.is_some_and(|x| x == target_id)
                || explicit.is_none() && owner.is_some_and(|x| x.id == target_id);
            let scoped_belongs = belongs && scope.is_none_or(|v| coordinates_in_scope(o, v));
            if let Some(formula) = o
                .get("formula")
                .and_then(Value::as_str)
                .map(str::to_owned)
            {
                if scoped_belongs {
                    let target = owner.cloned().unwrap_or_else(|| SheetIdentity {
                        id: target_id.into(),
                        name: target_id.into(),
                    });
                    let own = owner.cloned().unwrap_or_else(|| target.clone());
                    o.insert(
                        "formula".into(),
                        Value::String(references::remap_axis_in_scope(
                            &formula, &own, &target, axis, at, count, direction, scope,
                        )?),
                    );
                }
            }
            map_coordinates(o, scoped_belongs, axis, at, count, direction)?;
            let keys: Vec<String> = o.keys().cloned().collect();
            for k in keys {
                if matches!(
                    k.as_str(),
                    "row"
                        | "column"
                        | "startRow"
                        | "endRow"
                        | "startColumn"
                        | "endColumn"
                        | "formula"
                ) {
                    continue;
                }
                if let Some(child) = o.get_mut(&k) {
                    shift_value(child, owner, target_id, axis, at, count, direction, scope)?;
                }
            }
        }
        Value::Array(a) => {
            for child in a {
                shift_value(child, owner, target_id, axis, at, count, direction, scope)?;
            }
        }
        _ => {}
    }
    Ok(())
}
fn map_coordinates(
    o: &mut Map<String, Value>,
    belongs: bool,
    axis: Axis,
    at: u32,
    count: u32,
    d: Direction,
) -> KernelResult<()> {
    if !belongs {
        return Ok(());
    }
    let (key, end_key) = if axis == Axis::Row {
        ("row", "endRow")
    } else {
        ("column", "endColumn")
    };
    if let Some(v) = o.get(key).and_then(Value::as_u64) {
        o.insert(key.into(), Value::from(map_coord(v as u32, at, count, d)?));
    }
    if let Some(v) = o.get(end_key).and_then(Value::as_u64) {
        o.insert(
            end_key.into(),
            Value::from(map_coord(v as u32, at, count, d)?),
        );
    }
    for map_key in [
        if axis == Axis::Row {
            "hiddenRows"
        } else {
            "hiddenColumns"
        },
        if axis == Axis::Row {
            "rowHeightsPx"
        } else {
            "columnWidthsPx"
        },
    ] {
        if let Some(Value::Object(old)) = o.get(map_key) {
            let mut next = Map::new();
            for (k, v) in old {
                if let Ok(n) = k.parse::<u32>() {
                    if let Some(mapped) = shift_index(n, at, count, d) {
                        next.insert(mapped.to_string(), v.clone());
                    }
                } else {
                    next.insert(k.clone(), v.clone());
                }
            }
            o.insert(map_key.into(), Value::Object(next));
        }
    }
    Ok(())
}
fn coordinates_in_scope(o: &Map<String, Value>, scope: &RangeRef) -> bool {
    if let (Some(row), Some(column)) = (
        o.get("row").and_then(Value::as_u64),
        o.get("column").and_then(Value::as_u64),
    ) {
        return scope.contains(&CellAddress {
            sheet_id: scope.sheet_id.clone(),
            row: row as u32,
            column: column as u32,
        });
    }
    if let (Some(start_row), Some(end_row), Some(start_column), Some(end_column)) = (
        o.get("startRow").and_then(Value::as_u64),
        o.get("endRow").and_then(Value::as_u64),
        o.get("startColumn").and_then(Value::as_u64),
        o.get("endColumn").and_then(Value::as_u64),
    ) {
        return start_row as u32 <= scope.end_row
            && end_row as u32 >= scope.start_row
            && start_column as u32 <= scope.end_column
            && end_column as u32 >= scope.start_column;
    }
    false
}
fn map_coord(v: u32, at: u32, count: u32, d: Direction) -> KernelResult<u32> {
    shift_index(v, at, count, d)
        .ok_or_else(|| invalid("Structural delete would lose an anchored object"))
}

fn rewrite_rule_formulas(v: &mut Value, old: &str, new: &str) {
    match v {
        Value::Object(o) => {
            for k in ["formula", "formula1", "formula2", "value1", "value2"] {
                if let Some(Value::String(s)) = o.get_mut(k) {
                    if let Ok(next) = references::rename_sheet(s, old, new) {
                        *s = next;
                    }
                }
            }
            for child in o.values_mut() {
                rewrite_rule_formulas(child, old, new);
            }
        }
        Value::Array(a) => {
            for child in a {
                rewrite_rule_formulas(child, old, new);
            }
        }
        _ => {}
    }
}
fn remap_sheet_refs(
    v: &mut Value,
    source_id: &str,
    target_id: &str,
    source_name: &str,
    target_name: &str,
) {
    match v {
        Value::Object(o) => {
            if o.get("sheetId").and_then(Value::as_str) == Some(source_id) {
                o.insert("sheetId".into(), Value::String(target_id.into()));
            }
            if let Some(Value::String(s)) = o.get_mut("formula") {
                if let Ok(n) = references::rename_sheet(s, source_name, target_name) {
                    *s = n;
                }
            }
            for child in o.values_mut() {
                remap_sheet_refs(child, source_id, target_id, source_name, target_name);
            }
        }
        Value::Array(a) => {
            for child in a {
                remap_sheet_refs(child, source_id, target_id, source_name, target_name);
            }
        }
        _ => {}
    }
}

fn validate_delete_metadata(
    tx: &Transaction,
    sheet_id: &str,
    axis: Axis,
    at: u32,
    count: u32,
) -> KernelResult<()> {
    let end = at + count - 1;
    let s = tx.sheet(sheet_id)?;
    for v in s.metadata.values() {
        reject_deleted_anchors(v, sheet_id, axis, at, end)?;
    }
    Ok(())
}
fn reject_deleted_anchors(
    v: &Value,
    sheet_id: &str,
    axis: Axis,
    at: u32,
    end: u32,
) -> KernelResult<()> {
    match v {
        Value::Object(o) => {
            if o.get("sheetId")
                .and_then(Value::as_str)
                .is_some_and(|x| x != sheet_id)
            {
                return Ok(());
            }
            let key = if axis == Axis::Row { "row" } else { "column" };
            if let Some(n) = o.get(key).and_then(Value::as_u64) {
                if (at..=end).contains(&(n as u32))
                    && (o.contains_key("anchor")
                        || o.contains_key("id")
                        || o.contains_key("row") && o.contains_key("column"))
                {
                    return Err(invalid("Structural delete would lose metadata anchor"));
                }
            }
            for child in o.values() {
                reject_deleted_anchors(child, sheet_id, axis, at, end)?;
            }
        }
        Value::Array(a) => {
            for child in a {
                reject_deleted_anchors(child, sheet_id, axis, at, end)?;
            }
        }
        _ => {}
    }
    Ok(())
}

fn permute_metadata(
    v: &mut Value,
    sheet_id: &str,
    range: &RangeRef,
    mapping: &[u32],
) -> KernelResult<()> {
    match v {
        Value::Object(o) => {
            let belongs = o
                .get("sheetId")
                .and_then(Value::as_str)
                .is_none_or(|x| x == sheet_id);
            if belongs {
                if let Some(row) = o.get("row").and_then(Value::as_u64) {
                    if (range.start_row..=range.end_row).contains(&(row as u32)) {
                        o.insert(
                            "row".into(),
                            Value::from(remap_row(row as u32, range, mapping)),
                        );
                    }
                }
                if let (Some(a), Some(b)) = (
                    o.get("startRow").and_then(Value::as_u64),
                    o.get("endRow").and_then(Value::as_u64),
                ) {
                    if a as u32 >= range.start_row && b as u32 <= range.end_row {
                        let x = remap_row(a as u32, range, mapping);
                        let y = remap_row(b as u32, range, mapping);
                        o.insert("startRow".into(), Value::from(x.min(y)));
                        o.insert("endRow".into(), Value::from(x.max(y)));
                    } else if a as u32 <= range.end_row && b as u32 >= range.start_row {
                        return Err(invalid(
                            "Row permutation cannot exactly remap metadata range",
                        ));
                    }
                }
            }
            for child in o.values_mut() {
                permute_metadata(child, sheet_id, range, mapping)?;
            }
        }
        Value::Array(a) => {
            for child in a {
                permute_metadata(child, sheet_id, range, mapping)?;
            }
        }
        _ => {}
    }
    Ok(())
}
fn remap_row(row: u32, range: &RangeRef, mapping: &[u32]) -> u32 {
    range.start_row
        + mapping
            .iter()
            .position(|x| *x == row)
            .unwrap_or((row - range.start_row) as usize) as u32
}
fn permute_manifest_metadata(
    tx: &mut Transaction,
    sheet_id: &str,
    range: &RangeRef,
    mapping: &[u32],
) -> KernelResult<()> {
    for v in tx.manifest.metadata.values_mut() {
        permute_metadata(v, sheet_id, range, mapping)?;
    }
    Ok(())
}

fn has_external_reference(
    tx: &Transaction,
    source_id: &str,
    source_name: &str,
) -> KernelResult<bool> {
    for s in tx
        .manifest
        .sheets
        .iter()
        .filter(|x| x.sheet_id != source_id)
    {
        for (_, c) in tx.cells(&s.sheet_id)? {
            if let Some(f) = c.formula {
                if references::rename_sheet(&f, source_name, &format!("{source_name}__deleted__"))?
                    != f
                {
                    return Ok(true);
                }
            }
        }
        for v in s.metadata.values() {
            if metadata_mentions(v, source_id, source_name) {
                return Ok(true);
            }
        }
    }
    Ok(false)
}
fn metadata_mentions(v: &Value, source_id: &str, source_name: &str) -> bool {
    match v {
        Value::Object(o) => {
            if o.get("sheetId").and_then(Value::as_str) == Some(source_id) {
                return true;
            }
            if o.get("formula").and_then(Value::as_str).is_some_and(|f| {
                references::rename_sheet(f, source_name, &format!("{source_name}__deleted__"))
                    .map(|x| x != f)
                    .unwrap_or(false)
            }) {
                return true;
            }
            o.values()
                .any(|x| metadata_mentions(x, source_id, source_name))
        }
        Value::Array(a) => a
            .iter()
            .any(|x| metadata_mentions(x, source_id, source_name)),
        _ => false,
    }
}
