//! Canonical cell and range mutations.
//!
//! This module is deliberately the semantic owner of cell values, cell
//! presentation metadata, templates, and fill operations.  The browser may
//! send a plan (in particular for undo), but an applied fill is calculated
//! from the transaction's source cells here; an `after` image from a client
//! is never used as the source of truth.

use crate::{Transaction, integer, invalid, range, text};
use kernel_core::{Cell, CellAddress, KernelResult, RangeRef, Scalar};
use serde_json::{Map, Value};
use std::collections::{BTreeMap, BTreeSet, HashMap};

const MAX_ROWS: u32 = 1_048_576;
const MAX_COLUMNS: u32 = 16_384;

/// Apply one canonical cell mutation.
pub fn apply(tx: &mut Transaction<'_>, id: &str, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    match id {
        "cell.set" => cell_set(tx, sheet_id, p),
        "cell.restore" => cell_restore(tx, sheet_id, p),
        "cell.editor.set" => editor_set(tx, sheet_id, p),
        "cellTemplate.set" => template_set(tx, p),
        "cellTemplate.remove" => template_remove(tx, p),
        "range.set" => range_set(tx, sheet_id, p),
        "range.paste" => range_paste(tx, sheet_id, p),
        "range.clear" => range_clear(tx, sheet_id, p),
        "range.clear.restore" => range_clear_restore(tx, sheet_id, p),
        "style.set" => style_set(tx, sheet_id, p, false),
        "style.preset.set" => style_set(tx, sheet_id, p, true),
        "fill.applied" => fill_applied(tx, sheet_id, p),
        "fill.restored" => fill_restored(tx, sheet_id, p),
        _ => Ok(false),
    }
}

fn cell_set(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let row = integer(p, "row")?;
    let column = integer(p, "column")?;
    bounds(tx, sheet_id, row, column)?;
    let value = p
        .as_object()
        .and_then(|object| object.get("value"))
        .ok_or_else(|| invalid("cell.set value is required"))?;
    let cell = parse_cell(value)?;
    tx.write(sheet_id, row, column, Some(cell))?;
    Ok(true)
}

fn cell_restore(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let row = integer(p, "row")?;
    let column = integer(p, "column")?;
    bounds(tx, sheet_id, row, column)?;
    let previous = p.as_object().and_then(|object| object.get("previous"));
    let cell = match previous {
        None | Some(Value::Null) => None,
        Some(value) => Some(parse_cell(value)?),
    };
    tx.write(sheet_id, row, column, cell)?;
    Ok(true)
}

fn range_set(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let start_row = integer(p, "startRow")?;
    let start_column = integer(p, "startColumn")?;
    let values = p
        .as_object()
        .and_then(|object| object.get("values"))
        .and_then(Value::as_array)
        .ok_or_else(|| invalid("range.set values must be an array"))?;
    let mut writes = Vec::new();
    for (row_offset, row) in values.iter().enumerate() {
        let row_values = row
            .as_array()
            .ok_or_else(|| invalid("range.set rows must be arrays"))?;
        for (column_offset, value) in row_values.iter().enumerate() {
            let row = offset(start_row, row_offset)?;
            let column = offset(start_column, column_offset)?;
            bounds(tx, sheet_id, row, column)?;
            if value.is_null() {
                continue;
            }
            writes.push((row, column, Some(parse_cell(value)?)));
        }
    }
    for (row, column, cell) in writes {
        tx.write(sheet_id, row, column, cell)?;
    }
    Ok(true)
}

fn range_clear(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let target = range(tx, sheet_id, required(p, "range")?)?;
    let family = text(p, "family")?;
    if !matches!(
        family,
        "all" | "contents" | "formats" | "comments-and-notes" | "hyperlinks"
    ) {
        return Err(invalid(format!("Unsupported clear family: {family}")));
    }
    let entries = tx.cells_in_range(&target)?;
    for (address, original) in entries {
        if !target.contains(&address) {
            continue;
        }
        let next = match family {
            "all" => None,
            "contents" => Some(clear_contents(original)),
            "formats" => Some(clear_formats(original)),
            "comments-and-notes" | "hyperlinks" => Some(clear_metadata_family(original, family)),
            _ => unreachable!(),
        };
        if family == "comments-and-notes" || family == "hyperlinks" {
            if let Some(cell) = next {
                tx.write(sheet_id, address.row, address.column, Some(cell))?;
            }
        } else {
            tx.write(sheet_id, address.row, address.column, next)?;
        }
    }
    if family == "comments-and-notes" || family == "all" {
        if let Some(existing) = tx.sheet(sheet_id)?.metadata.get("review").cloned() {
            let mut review = existing;
            clear_review_metadata_value(&mut review, &target);
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert("review".into(), review);
        }
    }
    if family == "hyperlinks" || family == "all" {
        if let Some(existing) = tx.sheet(sheet_id)?.metadata.get("hyperlinks").cloned() {
            if !existing.is_array() {
                return Err(invalid("Worksheet hyperlinks metadata is malformed"));
            }
            let mut links = existing;
            clear_hyperlinks_value(&mut links, &target);
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert("hyperlinks".into(), links);
        }
    }
    if family == "formats" || family == "all" {
        for key in ["conditionalFormats", "dataValidations"] {
            let current = match tx.sheet(sheet_id)?.metadata.get(key) {
                None => Vec::new(),
                Some(value) => value
                    .as_array()
                    .cloned()
                    .ok_or_else(|| invalid(format!("Worksheet metadata {key} must be an array")))?,
            };
            let next: Vec<Value> = current
                .into_iter()
                .flat_map(|rule| subtract_rule(rule, &target))
                .collect();
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert(key.into(), Value::Array(next));
        }
    }
    Ok(true)
}

fn range_clear_restore(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let target = range(tx, sheet_id, required(p, "range")?)?;
    let snapshot = required(p, "snapshot")?;
    let entries = tx.cells_in_range(&target)?;
    for (address, _) in entries {
        if target.contains(&address) {
            tx.write(sheet_id, address.row, address.column, None)?;
        }
    }
    let cells = snapshot
        .as_object()
        .and_then(|object| object.get("cells"))
        .and_then(Value::as_array)
        .ok_or_else(|| invalid("range.clear.restore cells must be an array"))?;
    for entry in cells {
        let row = integer(entry, "row")?;
        let column = integer(entry, "column")?;
        bounds(tx, sheet_id, row, column)?;
        if !target.contains(&CellAddress {
            sheet_id: sheet_id.to_owned(),
            row,
            column,
        }) {
            return Err(invalid("Range restore cell is outside its range"));
        }
        let value = entry.as_object().and_then(|object| object.get("value"));
        let cell = match value {
            None | Some(Value::Null) => None,
            Some(value) => Some(parse_cell(value)?),
        };
        tx.write(sheet_id, row, column, cell)?;
    }
    if let Some(review) = snapshot.get("review") {
        let mut current = tx.sheet(sheet_id)?.metadata.get("review").cloned().unwrap_or_else(|| serde_json::json!({"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}));
        clear_review_metadata_value(&mut current, &target);
        copy_review_metadata(&mut current, review, &target, &target)?;
        tx.sheet_mut(sheet_id)?
            .metadata
            .insert("review".into(), current);
    } else if snapshot.get("notes").is_some() || snapshot.get("comments").is_some() {
        let review = review_from_snapshot(snapshot)?;
        let mut current = tx.sheet(sheet_id)?.metadata.get("review").cloned().unwrap_or_else(|| serde_json::json!({"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}));
        clear_review_metadata_value(&mut current, &target);
        copy_review_metadata(&mut current, &review, &target, &target)?;
        tx.sheet_mut(sheet_id)?
            .metadata
            .insert("review".into(), current);
    }
    if let Some(links) = snapshot.get("hyperlinks") {
        if !links.is_array() {
            return Err(invalid("Range restore hyperlinks must be an array"));
        }
        let mut next = tx
            .sheet(sheet_id)?
            .metadata
            .get("hyperlinks")
            .cloned()
            .unwrap_or(Value::Array(Vec::new()));
        if !next.is_array() {
            return Err(invalid("Worksheet hyperlinks metadata is malformed"));
        }
        clear_hyperlinks_value(&mut next, &target);
        copy_hyperlinks(&mut next, links, &target, &target, false)?;
        tx.sheet_mut(sheet_id)?
            .metadata
            .insert("hyperlinks".into(), next);
    }
    for (snapshot_key, metadata_key) in [
        ("conditionalFormats", "conditionalFormats"),
        ("dataValidations", "dataValidations"),
    ] {
        if let Some(value) = snapshot.get(snapshot_key) {
            if !value.is_array() {
                return Err(invalid(format!(
                    "Range restore {snapshot_key} must be an array"
                )));
            }
            copy_rules(tx, sheet_id, &target, &target, Some(value), metadata_key)?;
        }
    }
    Ok(true)
}

fn range_paste(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    // The source worksheet is authoritative for internal copy/cut payloads;
    // occupiedCells are used only for explicitly external clipboard input.
    canonical_paste(tx, sheet_id, p)?;
    Ok(true)
}

fn canonical_paste(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let clipboard = required(p, "clipboard")?;
    let source = any_range(tx, required(clipboard, "range")?)?;
    let extent = required(clipboard, "sourceExtent")?;
    let source_rows = integer(extent, "rows")?;
    let source_columns = integer(extent, "columns")?;
    if source_rows == 0
        || source_columns == 0
        || source.end_row - source.start_row + 1 != source_rows
        || source.end_column - source.start_column + 1 != source_columns
    {
        return Err(invalid(
            "Clipboard source extent does not match sourceRange",
        ));
    }
    let origin = required(p, "targetOrigin")?;
    let target_row = integer(origin, "row")?;
    let target_column = integer(origin, "column")?;
    let spec = required(p, "spec")?;
    let content = text(spec, "content")?;
    let formatting = text(spec, "formatting")?;
    let operation = text(spec, "operation")?;
    let transpose = spec
        .get("transpose")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    let skip_blanks = spec
        .get("skipBlanks")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    let link = spec.get("link").and_then(Value::as_bool).unwrap_or(false);
    if !matches!(content, "none" | "all" | "values" | "formulas")
        || !matches!(
            formatting,
            "all"
                | "none"
                | "number-format"
                | "source-formatting"
                | "all-except-borders"
                | "source-theme"
        )
        || !matches!(
            operation,
            "none" | "add" | "subtract" | "multiply" | "divide"
        )
    {
        return Err(invalid("Paste specification is invalid"));
    }
    if link && operation != "none" {
        return Err(invalid("Paste link cannot combine with arithmetic"));
    }
    let target_rows = if transpose {
        source_columns
    } else {
        source_rows
    };
    let target_columns = if transpose {
        source_rows
    } else {
        source_columns
    };
    let target = RangeRef {
        sheet_id: sheet_id.to_owned(),
        start_row: target_row,
        end_row: target_row
            .checked_add(target_rows - 1)
            .ok_or_else(|| invalid("Paste target overflows worksheet"))?,
        start_column: target_column,
        end_column: target_column
            .checked_add(target_columns - 1)
            .ok_or_else(|| invalid("Paste target overflows worksheet"))?,
    };
    target.validate()?;
    let target_sheet = tx.sheet(sheet_id)?;
    if target.end_row >= target_sheet.row_count || target.end_column >= target_sheet.column_count {
        return Err(invalid("Paste target is outside worksheet bounds"));
    }
    let transfer = text(p, "transfer")?;
    if !matches!(transfer, "copy" | "move") {
        return Err(invalid("Paste transfer is invalid"));
    }
    let clear_source = transfer == "move";
    if clear_source && ranges_intersect(&source, &target) {
        return Err(invalid("Cut source and target ranges may not overlap"));
    }

    let external = clipboard.get("source").and_then(Value::as_str) == Some("external");
    if external && formatting == "source-theme" {
        return Err(kernel_core::KernelError::new(
            "UNSUPPORTED_FEATURE",
            "External source-theme paste requires a canonical theme resource mapping",
        )
        .recover("paste-source-formatting"));
    }
    let occupied = clipboard
        .get("occupiedCells")
        .and_then(Value::as_array)
        .ok_or_else(|| invalid("Clipboard occupiedCells must be an array"))?;
    let mut external_cells = HashMap::new();
    if external {
        for entry in occupied {
            let row_offset = integer(entry, "rowOffset")?;
            let column_offset = integer(entry, "columnOffset")?;
            if row_offset >= source_rows || column_offset >= source_columns {
                return Err(invalid("Clipboard occupied cell is outside sourceExtent"));
            }
            let cell = parse_cell(required(entry, "value")?)?;
            if external_cells
                .insert((row_offset, column_offset), cell)
                .is_some()
            {
                return Err(invalid(
                    "Clipboard occupied cells contain a duplicate coordinate",
                ));
            }
        }
    }
    let mut writes: Vec<(String, u32, u32, Option<Cell>)> = Vec::new();
    for source_row_offset in 0..source_rows {
        for source_column_offset in 0..source_columns {
            let source_row = source.start_row + source_row_offset;
            let source_column = source.start_column + source_column_offset;
            let source_cell = if external {
                external_cells
                    .get(&(source_row_offset, source_column_offset))
                    .cloned()
            } else {
                tx.read(&source.sheet_id, source_row, source_column)?
            };
            let row = target.start_row
                + if transpose {
                    source_column_offset
                } else {
                    source_row_offset
                };
            let column = target.start_column
                + if transpose {
                    source_row_offset
                } else {
                    source_column_offset
                };
            let current = tx.read(sheet_id, row, column)?;
            let next = paste_cell(
                spec,
                transfer,
                source_cell,
                current.clone(),
                row as i64 - source_row as i64,
                column as i64 - source_column as i64,
                &source,
                &tx.sheet(&source.sheet_id)?.name,
                source_row_offset,
                source_column_offset,
            )?;
            if next.is_some() && next != current {
                writes.push((sheet_id.to_owned(), row, column, next));
            }
        }
    }
    // A move clears the source after all source reads have completed.  This
    // is deliberately separate from target writes for cut semantics.
    if clear_source {
        for (row, column) in coordinates(&source) {
            tx.write(&source.sheet_id, row, column, None)?;
        }
    }
    for (_, row, column, cell) in writes {
        tx.write(sheet_id, row, column, cell)?;
    }
    paste_metadata(
        tx,
        sheet_id,
        &source,
        &target,
        spec,
        clipboard,
        clear_source,
    )?;
    Ok(())
}

fn paste_cell(
    spec: &Value,
    transfer: &str,
    source: Option<Cell>,
    target: Option<Cell>,
    row_delta: i64,
    column_delta: i64,
    source_range: &RangeRef,
    source_sheet_name: &str,
    source_row_offset: u32,
    source_column_offset: u32,
) -> KernelResult<Option<Cell>> {
    let content = text(spec, "content")?;
    let formatting = text(spec, "formatting")?;
    let operation = text(spec, "operation")?;
    let skip_blanks = spec
        .get("skipBlanks")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    let link = spec.get("link").and_then(Value::as_bool).unwrap_or(false);
    let Some(source) = source else {
        return if skip_blanks {
            Ok(None)
        } else {
            Ok(Some(Cell::default()))
        };
    };
    let source_blank = matches!(source.value, Scalar::Null) && source.formula.is_none();
    if skip_blanks && source_blank {
        return Ok(None);
    }
    let mut destination = target.clone().unwrap_or_default();
    if link {
        destination.value = Scalar::Null;
        destination.formula = Some(format!(
            "='{}'!{}{}",
            source_sheet_name.replace('\'', "''"),
            column_label(source_range.start_column + source_column_offset),
            source_range.start_row + source_row_offset + 1
        ));
        for key in ["formulaValue", "displayValue", "formulaMetadata"] {
            destination.metadata.remove(key);
        }
        return Ok(Some(destination));
    }
    let source_formula = if let Some(formula) = source.formula.clone() {
        if transfer == "move" {
            Some(formula)
        } else {
            Some(kernel_formula::references::offset(
                &formula,
                i32::try_from(row_delta)
                    .map_err(|_| invalid("Paste row offset is outside supported bounds"))?,
                i32::try_from(column_delta)
                    .map_err(|_| invalid("Paste column offset is outside supported bounds"))?,
            )?)
        }
    } else {
        None
    };
    if operation != "none" {
        if source_formula.is_some() || target.as_ref().is_some_and(|cell| cell.formula.is_some()) {
            return Err(invalid("Paste arithmetic cannot operate on formula cells"));
        }
        let right = match source.value {
            Scalar::Number(value) if value.is_finite() => value,
            _ => {
                return Err(invalid(
                    "Paste arithmetic requires numeric source and target values",
                ));
            }
        };
        let left = match target.as_ref().map(|cell| &cell.value) {
            None | Some(Scalar::Null) => 0.0,
            Some(Scalar::Number(value)) if value.is_finite() => *value,
            _ => {
                return Err(invalid(
                    "Paste arithmetic requires numeric source and target values",
                ));
            }
        };
        let value = match operation {
            "add" => left + right,
            "subtract" => left - right,
            "multiply" => left * right,
            "divide" if right != 0.0 => left / right,
            "divide" => return Err(invalid("Paste arithmetic divide cannot use zero")),
            _ => return Err(invalid("Paste arithmetic operation is invalid")),
        };
        if !value.is_finite() {
            return Err(invalid("Paste arithmetic result is not finite"));
        }
        let mut next = clear_formula_provenance(destination);
        next.value = Scalar::Number(value);
        next.formula = None;
        return Ok(Some(next));
    }
    if content == "none" {
        if formatting == "none" {
            return Ok(None);
        }
        let mut next = destination;
        if matches!(formatting, "all" | "source-formatting" | "source-theme") {
            next.metadata.insert(
                "style".into(),
                source.metadata.get("style").cloned().unwrap_or(Value::Null),
            );
        }
        if matches!(
            formatting,
            "all" | "source-formatting" | "source-theme" | "number-format"
        ) {
            copy_metadata_key(&mut next, &source, "numberFormat");
        }
        return Ok(Some(next));
    }
    if content == "values" {
        let mut next = Cell {
            value: source.value.clone(),
            formula: None,
            metadata: BTreeMap::new(),
        };
        if matches!(
            formatting,
            "source-formatting" | "all" | "source-theme" | "all-except-borders"
        ) {
            if let Some(style) = source.metadata.get("style") {
                next.metadata.insert("style".into(), style.clone());
            }
        }
        if formatting == "number-format"
            || matches!(formatting, "source-formatting" | "all" | "source-theme")
        {
            copy_metadata_key(&mut next, &source, "numberFormat");
        }
        if formatting == "all-except-borders" {
            preserve_borders(&mut next, target.as_ref());
        }
        return Ok(Some(next));
    }
    if content == "formulas" {
        let mut next = clear_formula_provenance(destination);
        if let Some(formula) = source_formula {
            next.value = Scalar::Null;
            next.formula = Some(formula);
        } else {
            next.value = source.value.clone();
            next.formula = None;
        }
        apply_paste_formatting(&mut next, &source, target.as_ref(), formatting);
        return Ok(Some(next));
    }
    if formatting == "none" {
        let mut next = clear_formula_provenance(destination);
        next.value = source.value.clone();
        next.formula = source_formula;
        return Ok(Some(next));
    }
    if formatting == "number-format" {
        let mut next = clear_formula_provenance(destination);
        next.value = source.value.clone();
        next.formula = source_formula;
        copy_metadata_key(&mut next, &source, "numberFormat");
        return Ok(Some(next));
    }
    let mut next = clear_formula_provenance(source);
    next.formula = source_formula;
    if formatting == "all-except-borders" {
        preserve_borders(&mut next, target.as_ref());
    }
    Ok(Some(next))
}

fn apply_paste_formatting(next: &mut Cell, source: &Cell, target: Option<&Cell>, formatting: &str) {
    match formatting {
        "number-format" => copy_metadata_key(next, source, "numberFormat"),
        "source-formatting" | "all" | "source-theme" => {
            if let Some(style) = source.metadata.get("style") {
                next.metadata.insert("style".into(), style.clone());
            }
            copy_metadata_key(next, source, "numberFormat");
        }
        "all-except-borders" => {
            if let Some(style) = source.metadata.get("style") {
                next.metadata.insert("style".into(), style.clone());
            }
            preserve_borders(next, target);
        }
        _ => {}
    }
}

fn copy_metadata_key(destination: &mut Cell, source: &Cell, key: &str) {
    if let Some(value) = source.metadata.get(key) {
        destination.metadata.insert(key.into(), value.clone());
    } else {
        destination.metadata.remove(key);
    }
}
fn preserve_borders(destination: &mut Cell, target: Option<&Cell>) {
    if let Some(Value::Object(style)) = destination.metadata.get_mut("style") {
        let borders = target
            .and_then(|cell| cell.metadata.get("style"))
            .and_then(|value| value.get("borders"))
            .cloned();
        if let Some(borders) = borders {
            style.insert("borders".into(), borders);
        } else {
            style.remove("borders");
        }
    }
}
fn clear_formula_provenance(mut cell: Cell) -> Cell {
    for key in ["formulaValue", "displayValue", "formulaMetadata"] {
        cell.metadata.remove(key);
    }
    cell
}
fn ranges_intersect(left: &RangeRef, right: &RangeRef) -> bool {
    left.sheet_id == right.sheet_id
        && left.start_row <= right.end_row
        && left.end_row >= right.start_row
        && left.start_column <= right.end_column
        && left.end_column >= right.start_column
}
fn column_label(mut value: u32) -> String {
    let mut result = String::new();
    value += 1;
    while value > 0 {
        let remainder = ((value - 1) % 26) as u8;
        result.insert(0, (b'A' + remainder) as char);
        value = (value - 1) / 26;
    }
    result
}

fn paste_metadata(
    tx: &mut Transaction<'_>,
    target_sheet_id: &str,
    source: &RangeRef,
    target: &RangeRef,
    spec: &Value,
    clipboard: &Value,
    moved: bool,
) -> KernelResult<()> {
    let metadata = spec
        .get("metadata")
        .ok_or_else(|| invalid("Paste metadata specification is required"))?;
    let mut source_metadata = clipboard
        .get("rangeMetadata")
        .cloned()
        .unwrap_or_else(|| Value::Object(Map::new()));
    if clipboard.get("source").and_then(Value::as_str) == Some("external") {
        external_review_metadata(&mut source_metadata, source)?;
        external_hyperlink_metadata(&mut source_metadata, source)?;
    }
    let source_sheet = tx.sheet(source.sheet_id.as_str())?.metadata.clone();
    let include_comments = metadata
        .get("commentsNotes")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    let include_links = metadata
        .get("hyperlinks")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    let include_validations = metadata
        .get("validation")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    let include_cf = metadata
        .get("conditionalFormats")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    let include_widths = metadata
        .get("columnWidths")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    if include_comments {
        let mut review = tx.sheet(target_sheet_id)?.metadata.get("review").cloned().unwrap_or_else(|| serde_json::json!({"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}));
        clear_review_metadata_value(&mut review, target);
        let source_review = source_sheet.get("review").cloned().unwrap_or_else(|| {
            source_metadata
                .get("review")
                .cloned()
                .unwrap_or(Value::Null)
        });
        copy_review_metadata(&mut review, &source_review, source, target)?;
        tx.sheet_mut(target_sheet_id)?
            .metadata
            .insert("review".into(), review);
        if moved {
            let mut source_review = tx
                .sheet(source.sheet_id.as_str())?
                .metadata
                .get("review")
                .cloned()
                .unwrap_or(Value::Null);
            clear_review_metadata_value(&mut source_review, source);
            tx.sheet_mut(&source.sheet_id)?
                .metadata
                .insert("review".into(), source_review);
        }
    }
    if include_links {
        let mut links = tx
            .sheet(target_sheet_id)?
            .metadata
            .get("hyperlinks")
            .cloned()
            .unwrap_or_else(|| Value::Array(Vec::new()));
        if !links.is_array() {
            return Err(invalid("Worksheet hyperlinks metadata is malformed"));
        }
        clear_hyperlinks_value(&mut links, target);
        let source_links = source_sheet
            .get("hyperlinks")
            .cloned()
            .or_else(|| source_metadata.get("hyperlinks").cloned())
            .unwrap_or(Value::Array(Vec::new()));
        copy_hyperlinks(
            &mut links,
            &source_links,
            source,
            target,
            spec.get("transpose")
                .and_then(Value::as_bool)
                .unwrap_or(false),
        )?;
        tx.sheet_mut(target_sheet_id)?
            .metadata
            .insert("hyperlinks".into(), links);
        if moved {
            let mut source_links = tx
                .sheet(source.sheet_id.as_str())?
                .metadata
                .get("hyperlinks")
                .cloned()
                .unwrap_or(Value::Array(Vec::new()));
            if !source_links.is_array() {
                return Err(invalid("Source hyperlinks metadata is malformed"));
            }
            clear_hyperlinks_value(&mut source_links, source);
            tx.sheet_mut(&source.sheet_id)?
                .metadata
                .insert("hyperlinks".into(), source_links);
        }
    }
    if include_validations {
        copy_rules(
            tx,
            target_sheet_id,
            source,
            target,
            source_sheet
                .get("dataValidations")
                .or_else(|| source_metadata.get("validations")),
            "dataValidations",
        )?;
    }
    if include_cf {
        copy_rules(
            tx,
            target_sheet_id,
            source,
            target,
            source_sheet
                .get("conditionalFormats")
                .or_else(|| source_metadata.get("conditionalFormats")),
            "conditionalFormats",
        )?;
    }
    if include_widths {
        let widths = source_sheet
            .get("columnWidthsPx")
            .or_else(|| source_metadata.get("columnWidths"));
        if let Some(widths) = widths {
            copy_column_widths(tx, target_sheet_id, source, target, widths, moved)?;
            if moved {
                let mut source_widths = tx
                    .sheet(source.sheet_id.as_str())?
                    .metadata
                    .get("columnWidthsPx")
                    .cloned()
                    .unwrap_or(Value::Object(Map::new()));
                if let Some(object) = source_widths.as_object_mut() {
                    for column in source.start_column..=source.end_column {
                        object.remove(&column.to_string());
                    }
                }
                tx.sheet_mut(&source.sheet_id)?
                    .metadata
                    .insert("columnWidthsPx".into(), source_widths);
            }
        }
    }
    Ok(())
}

fn copy_review_metadata(
    target: &mut Value,
    source: &Value,
    source_range: &RangeRef,
    target_range: &RangeRef,
) -> KernelResult<()> {
    let target_object = target
        .as_object_mut()
        .ok_or_else(|| invalid("Review metadata is malformed"))?;
    for key in ["notesByCell", "notesById", "threadIdsByCell", "threadsById"] {
        if !target_object.get(key).is_some_and(Value::is_object) {
            target_object.insert(key.into(), Value::Object(Map::new()));
        }
    }
    let source_object = source
        .as_object()
        .ok_or_else(|| invalid("Source review metadata is malformed"))?;
    let notes_by_cell = source_object
        .get("notesByCell")
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    let notes_by_id = source_object
        .get("notesById")
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    let threads_by_cell = source_object
        .get("threadIdsByCell")
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    let threads_by_id = source_object
        .get("threadsById")
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    let mut target_notes_cell = target_object
        .get("notesByCell")
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    let mut target_notes_id = target_object
        .get("notesById")
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    let mut target_threads_cell = target_object
        .get("threadIdsByCell")
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    let mut target_threads_id = target_object
        .get("threadsById")
        .and_then(Value::as_object)
        .cloned()
        .unwrap_or_default();
    for (key, note_id) in notes_by_cell {
        let Some((row, column)) = parse_coordinate(&key) else {
            continue;
        };
        if !source_range.contains(&CellAddress {
            sheet_id: source_range.sheet_id.clone(),
            row,
            column,
        }) {
            continue;
        }
        let Some(note) = note_id.as_str().and_then(|id| notes_by_id.get(id)) else {
            continue;
        };
        let new_row = target_range.start_row + row - source_range.start_row;
        let new_column = target_range.start_column + column - source_range.start_column;
        let new_id = format!(
            "{}@paste:{}:{}",
            note_id.as_str().unwrap_or("note"),
            new_row,
            new_column
        );
        target_notes_cell.insert(
            format!("{}:{}", new_row, new_column),
            Value::String(new_id.clone()),
        );
        target_notes_id.insert(new_id, note.clone());
    }
    for (key, thread_ids) in threads_by_cell {
        let Some((row, column)) = parse_coordinate(&key) else {
            continue;
        };
        if !source_range.contains(&CellAddress {
            sheet_id: source_range.sheet_id.clone(),
            row,
            column,
        }) {
            continue;
        }
        let Some(ids) = thread_ids.as_array() else {
            continue;
        };
        let new_row = target_range.start_row + row - source_range.start_row;
        let new_column = target_range.start_column + column - source_range.start_column;
        let target_key = format!("{}:{}", new_row, new_column);
        let target_ids = target_threads_cell
            .entry(target_key)
            .or_insert_with(|| Value::Array(Vec::new()))
            .as_array_mut()
            .ok_or_else(|| invalid("Target review thread index is malformed"))?;
        for thread_id in ids {
            let Some(thread_id) = thread_id.as_str() else {
                continue;
            };
            let Some(thread) = threads_by_id.get(thread_id) else {
                continue;
            };
            let new_id = format!("{}@paste:{}:{}", thread_id, new_row, new_column);
            let mut next = thread.clone();
            if let Some(object) = next.as_object_mut() {
                object.insert("id".into(), Value::String(new_id.clone()));
                object.insert(
                    "sheetId".into(),
                    Value::String(target_range.sheet_id.clone()),
                );
                object.insert("row".into(), Value::from(new_row));
                object.insert("column".into(), Value::from(new_column));
            }
            target_ids.push(Value::String(new_id.clone()));
            target_threads_id.insert(new_id, next);
        }
    }
    target_object.insert("notesByCell".into(), Value::Object(target_notes_cell));
    target_object.insert("notesById".into(), Value::Object(target_notes_id));
    target_object.insert("threadIdsByCell".into(), Value::Object(target_threads_cell));
    target_object.insert("threadsById".into(), Value::Object(target_threads_id));
    Ok(())
}

fn external_review_metadata(metadata: &mut Value, source: &RangeRef) -> KernelResult<()> {
    let mut review =
        serde_json::json!({"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}});
    if let Some(notes) = metadata.get("notes").and_then(Value::as_array) {
        for entry in notes {
            let row = source.start_row + integer(entry, "rowOffset")?;
            let column = source.start_column + integer(entry, "columnOffset")?;
            let note = required(entry, "value")?;
            let id = text(note, "id")?;
            review["notesByCell"][format!("{row}:{column}")] = Value::String(id.into());
            review["notesById"][id] = note.clone();
        }
    }
    if let Some(comments) = metadata.get("comments").and_then(Value::as_array) {
        for entry in comments {
            let row = source.start_row + integer(entry, "rowOffset")?;
            let column = source.start_column + integer(entry, "columnOffset")?;
            let comment = required(entry, "value")?;
            let id = text(comment, "id")?;
            review["threadIdsByCell"][format!("{row}:{column}")] = serde_json::json!([id]);
            review["threadsById"][id] = comment.clone();
        }
    }
    metadata
        .as_object_mut()
        .ok_or_else(|| invalid("Clipboard rangeMetadata is malformed"))?
        .insert("review".into(), review);
    Ok(())
}

fn external_hyperlink_metadata(metadata: &mut Value, source: &RangeRef) -> KernelResult<()> {
    let mut links = Vec::new();
    if let Some(values) = metadata.get("hyperlinks").and_then(Value::as_array) {
        for entry in values {
            let row = source.start_row + integer(entry, "rowOffset")?;
            let column = source.start_column + integer(entry, "columnOffset")?;
            let value = required(entry, "value")?;
            links.push(serde_json::json!({"row":row,"column":column,"hyperlink":value}));
        }
    }
    metadata
        .as_object_mut()
        .ok_or_else(|| invalid("Clipboard rangeMetadata is malformed"))?
        .insert("hyperlinks".into(), Value::Array(links));
    Ok(())
}

fn clear_review_metadata_value(review: &mut Value, range: &RangeRef) {
    let Some(object) = review.as_object_mut() else {
        return;
    };
    let keys: Vec<String> = object
        .get("notesByCell")
        .and_then(Value::as_object)
        .map(|values| values.keys().cloned().collect())
        .unwrap_or_default();
    let mut note_ids = Vec::new();
    if let Some(values) = object.get_mut("notesByCell").and_then(Value::as_object_mut) {
        for key in keys {
            if parse_coordinate(&key).is_some_and(|(row, column)| {
                range.contains(&CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row,
                    column,
                })
            }) {
                if let Some(id) = values
                    .remove(&key)
                    .and_then(|value| value.as_str().map(str::to_owned))
                {
                    note_ids.push(id);
                }
            }
        }
    }
    if let Some(values) = object.get_mut("notesById").and_then(Value::as_object_mut) {
        for id in note_ids {
            values.remove(&id);
        }
    }
    let thread_keys: Vec<String> = object
        .get("threadIdsByCell")
        .and_then(Value::as_object)
        .map(|values| values.keys().cloned().collect())
        .unwrap_or_default();
    let mut thread_ids = Vec::new();
    if let Some(values) = object
        .get_mut("threadIdsByCell")
        .and_then(Value::as_object_mut)
    {
        for key in thread_keys {
            if parse_coordinate(&key).is_some_and(|(row, column)| {
                range.contains(&CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row,
                    column,
                })
            }) {
                if let Some(ids) = values
                    .remove(&key)
                    .and_then(|value| value.as_array().cloned())
                {
                    thread_ids.extend(
                        ids.into_iter()
                            .filter_map(|id| id.as_str().map(str::to_owned)),
                    );
                }
            }
        }
    }
    if let Some(values) = object.get_mut("threadsById").and_then(Value::as_object_mut) {
        for id in thread_ids {
            values.remove(&id);
        }
    }
}

fn copy_hyperlinks(
    target: &mut Value,
    source: &Value,
    source_range: &RangeRef,
    target_range: &RangeRef,
    transpose: bool,
) -> KernelResult<()> {
    let target_values = target
        .as_array_mut()
        .ok_or_else(|| invalid("Worksheet hyperlinks metadata is malformed"))?;
    target_values.retain(|value| {
        !value
            .get("row")
            .and_then(Value::as_u64)
            .zip(value.get("column").and_then(Value::as_u64))
            .is_some_and(|(row, column)| {
                target_range.start_row <= row as u32
                    && row as u32 <= target_range.end_row
                    && target_range.start_column <= column as u32
                    && column as u32 <= target_range.end_column
            })
    });
    let source_values = source
        .as_array()
        .cloned()
        .ok_or_else(|| invalid("Source hyperlinks metadata must be an array"))?;
    for value in source_values {
        let Some(row) = value
            .get("row")
            .and_then(Value::as_u64)
            .map(|value| value as u32)
        else {
            continue;
        };
        let Some(column) = value
            .get("column")
            .and_then(Value::as_u64)
            .map(|value| value as u32)
        else {
            continue;
        };
        if row < source_range.start_row
            || row > source_range.end_row
            || column < source_range.start_column
            || column > source_range.end_column
        {
            continue;
        }
        let row_offset = row - source_range.start_row;
        let column_offset = column - source_range.start_column;
        let target_row =
            target_range.start_row + if transpose { column_offset } else { row_offset };
        let target_column =
            target_range.start_column + if transpose { row_offset } else { column_offset };
        let mut next = value.clone();
        if let Some(object) = next.as_object_mut() {
            object.insert("row".into(), Value::from(target_row));
            object.insert("column".into(), Value::from(target_column));
        }
        target_values.push(next);
    }
    Ok(())
}

fn clear_hyperlinks_value(value: &mut Value, range: &RangeRef) {
    if let Some(values) = value.as_array_mut() {
        values.retain(|entry| {
            !(entry
                .get("row")
                .and_then(Value::as_u64)
                .is_some_and(|row| row as u32 >= range.start_row && row as u32 <= range.end_row)
                && entry
                    .get("column")
                    .and_then(Value::as_u64)
                    .is_some_and(|column| {
                        column as u32 >= range.start_column && column as u32 <= range.end_column
                    }))
        });
    }
}

fn review_from_snapshot(snapshot: &Value) -> KernelResult<Value> {
    let mut review =
        serde_json::json!({"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}});
    if let Some(notes) = snapshot.get("notes") {
        let notes = notes
            .as_array()
            .ok_or_else(|| invalid("Range restore notes must be an array"))?;
        for entry in notes {
            let row = integer(entry, "row")?;
            let column = integer(entry, "column")?;
            let note = entry
                .get("note")
                .or_else(|| entry.get("value"))
                .ok_or_else(|| invalid("Range restore note is missing"))?;
            let id = text(note, "id")?;
            review["notesByCell"][format!("{row}:{column}")] = Value::String(id.into());
            review["notesById"][id] = note.clone();
        }
    }
    if let Some(comments) = snapshot.get("comments") {
        let comments = comments
            .as_array()
            .ok_or_else(|| invalid("Range restore comments must be an array"))?;
        for comment in comments {
            let row = integer(comment, "row")?;
            let column = integer(comment, "column")?;
            let id = text(comment, "id")?;
            review["threadIdsByCell"][format!("{row}:{column}")] = serde_json::json!([id]);
            review["threadsById"][id] = comment.clone();
        }
    }
    Ok(review)
}

fn copy_rules(
    tx: &mut Transaction<'_>,
    target_sheet_id: &str,
    source: &RangeRef,
    target: &RangeRef,
    source_rules: Option<&Value>,
    key: &str,
) -> KernelResult<()> {
    let source_rules = match source_rules {
        None => Vec::new(),
        Some(value) => value
            .as_array()
            .cloned()
            .ok_or_else(|| invalid(format!("Source metadata {key} must be an array")))?,
    };
    let mut next_rules = match tx.sheet(target_sheet_id)?.metadata.get(key) {
        None => Vec::new(),
        Some(value) => value
            .as_array()
            .cloned()
            .ok_or_else(|| invalid(format!("Worksheet metadata {key} must be an array")))?,
    };
    next_rules = next_rules
        .into_iter()
        .flat_map(|rule| subtract_rule(rule, target))
        .collect();
    for rule in source_rules {
        if let Some(mapped) = map_rule(rule, source, target) {
            next_rules.push(mapped);
        }
    }
    tx.sheet_mut(target_sheet_id)?
        .metadata
        .insert(key.into(), Value::Array(next_rules));
    Ok(())
}

fn map_rule(mut rule: Value, source: &RangeRef, target: &RangeRef) -> Option<Value> {
    let ranges = rule.get("ranges")?.as_array()?.clone();
    let mut mapped = Vec::new();
    for range_value in ranges {
        let value = serde_json::from_value::<RangeRef>(range_value).ok()?;
        if !ranges_intersect(&value, source) {
            continue;
        }
        let clipped = intersect_range(&value, source)?;
        let mut result = clipped.clone();
        result.sheet_id = target.sheet_id.clone();
        result.start_row = target.start_row + clipped.start_row - source.start_row;
        result.end_row = target.start_row + clipped.end_row - source.start_row;
        result.start_column = target.start_column + clipped.start_column - source.start_column;
        result.end_column = target.start_column + clipped.end_column - source.start_column;
        mapped.push(serde_json::to_value(result).ok()?);
    }
    if mapped.is_empty() {
        return None;
    }
    rule.as_object_mut()?
        .insert("ranges".into(), Value::Array(mapped));
    Some(rule)
}

fn subtract_rule(mut rule: Value, cut: &RangeRef) -> Vec<Value> {
    let Some(ranges) = rule.get("ranges").and_then(Value::as_array) else {
        return vec![rule];
    };
    let mut output = Vec::new();
    for value in ranges {
        let Ok(base) = serde_json::from_value::<RangeRef>(value.clone()) else {
            continue;
        };
        if !ranges_intersect(&base, cut) {
            output.push(value.clone());
            continue;
        }
        for remainder in subtract_range(&base, cut) {
            if let Ok(value) = serde_json::to_value(remainder) {
                let mut copy = rule.clone();
                copy.as_object_mut()
                    .unwrap()
                    .insert("ranges".into(), Value::Array(vec![value]));
                output.push(copy);
            }
        }
    }
    output
}

fn subtract_range(base: &RangeRef, cut: &RangeRef) -> Vec<RangeRef> {
    let Some(overlap) = intersect_range(base, cut) else {
        return vec![base.clone()];
    };
    let mut out = Vec::new();
    if base.start_row < overlap.start_row {
        out.push(RangeRef {
            end_row: overlap.start_row - 1,
            ..base.clone()
        });
    }
    if overlap.end_row < base.end_row {
        out.push(RangeRef {
            start_row: overlap.end_row + 1,
            ..base.clone()
        });
    }
    if base.start_column < overlap.start_column {
        out.push(RangeRef {
            start_row: overlap.start_row,
            end_row: overlap.end_row,
            end_column: overlap.start_column - 1,
            ..base.clone()
        });
    }
    if overlap.end_column < base.end_column {
        out.push(RangeRef {
            start_row: overlap.start_row,
            end_row: overlap.end_row,
            start_column: overlap.end_column + 1,
            ..base.clone()
        });
    }
    out
}

fn intersect_range(left: &RangeRef, right: &RangeRef) -> Option<RangeRef> {
    if !ranges_intersect(left, right) {
        return None;
    }
    Some(RangeRef {
        sheet_id: left.sheet_id.clone(),
        start_row: left.start_row.max(right.start_row),
        end_row: left.end_row.min(right.end_row),
        start_column: left.start_column.max(right.start_column),
        end_column: left.end_column.min(right.end_column),
    })
}

fn copy_column_widths(
    tx: &mut Transaction<'_>,
    target_sheet_id: &str,
    source: &RangeRef,
    target: &RangeRef,
    widths: &Value,
    _moved: bool,
) -> KernelResult<()> {
    let mut destination = match tx.sheet(target_sheet_id)?.metadata.get("columnWidthsPx") {
        None => Map::new(),
        Some(value) => value
            .as_object()
            .cloned()
            .ok_or_else(|| invalid("Worksheet columnWidthsPx metadata is malformed"))?,
    };
    if let Some(object) = widths.as_object() {
        for (column, width) in object {
            if let Ok(column) = column.parse::<u32>() {
                if column >= source.start_column && column <= source.end_column {
                    destination.insert(
                        (target.start_column + column - source.start_column).to_string(),
                        width.clone(),
                    );
                }
            }
        }
    }
    if let Some(array) = widths.as_array() {
        for entry in array {
            let offset = integer(entry, "offset")?;
            let column = source.start_column + offset;
            if column <= source.end_column {
                if let Some(width) = entry.get("widthPx") {
                    destination.insert((target.start_column + offset).to_string(), width.clone());
                }
            }
        }
    } else if !widths.is_object() {
        return Err(invalid("Source column widths metadata is malformed"));
    }
    tx.sheet_mut(target_sheet_id)?
        .metadata
        .insert("columnWidthsPx".into(), Value::Object(destination));
    Ok(())
}

fn parse_coordinate(value: &str) -> Option<(u32, u32)> {
    let (row, column) = value.split_once(':')?;
    Some((row.parse().ok()?, column.parse().ok()?))
}

fn any_range(tx: &Transaction<'_>, value: &Value) -> KernelResult<RangeRef> {
    let result: RangeRef = serde_json::from_value(value.clone())
        .map_err(|error| invalid(format!("Invalid range: {error}")))?;
    result.validate()?;
    let sheet = tx.sheet(&result.sheet_id)?;
    if result.end_row >= sheet.row_count || result.end_column >= sheet.column_count {
        return Err(invalid("Range exceeds worksheet dimensions"));
    }
    Ok(result)
}

fn style_set(
    tx: &mut Transaction<'_>,
    sheet_id: &str,
    p: &Value,
    preset: bool,
) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let ranges = if let Some(value) = p.as_object().and_then(|object| object.get("ranges")) {
        value
            .as_array()
            .ok_or_else(|| invalid("style ranges must be an array"))?
            .iter()
            .map(|value| range(tx, sheet_id, value))
            .collect::<KernelResult<Vec<_>>>()?
    } else {
        vec![range(tx, sheet_id, required(p, "range")?)?]
    };
    let object = p
        .as_object()
        .ok_or_else(|| invalid("Style parameters must be an object"))?;
    let style = object.get("style").and_then(Value::as_object);
    let number_format = object.get("numberFormat").and_then(Value::as_str);
    let style_number_format = style
        .and_then(|value| value.get("numberFormat"))
        .and_then(Value::as_str);
    let clear_number_format = object
        .get("clearNumberFormat")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    if style.is_none() && number_format.is_none() && !clear_number_format {
        return Err(invalid("style.set requires style or numberFormat"));
    }
    let style_id = object.get("styleId").and_then(Value::as_str);
    for target in ranges {
        for (row, column) in coordinates(&target) {
            let mut cell = tx.read(sheet_id, row, column)?.unwrap_or_default();
            if let Some(style) = style {
                let replace = object
                    .get("replaceStyle")
                    .and_then(Value::as_bool)
                    .unwrap_or(preset);
                if replace {
                    cell.metadata.remove("style");
                }
                let mut merged = match cell.metadata.remove("style") {
                    Some(Value::Object(value)) if !replace => value,
                    _ => Map::new(),
                };
                for (key, value) in style {
                    merged.insert(key.clone(), value.clone());
                }
                if merged.is_empty() {
                    cell.metadata.remove("style");
                } else {
                    cell.metadata.insert("style".into(), Value::Object(merged));
                }
                cell.metadata.remove("displayValue");
            }
            if let Some(value) = number_format.or(style_number_format) {
                cell.metadata
                    .insert("numberFormat".into(), Value::String(value.into()));
            } else if clear_number_format {
                cell.metadata.remove("numberFormat");
            }
            if let Some(style_id) = style_id {
                cell.metadata
                    .insert("styleId".into(), Value::String(style_id.into()));
            }
            tx.write(sheet_id, row, column, Some(cell))?;
        }
    }
    Ok(true)
}

fn editor_set(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let ranges = p
        .as_object()
        .and_then(|object| object.get("ranges"))
        .and_then(Value::as_array)
        .ok_or_else(|| invalid("cell.editor.set ranges must be an array"))?;
    let editor = p.as_object().and_then(|object| object.get("editor"));
    for item in ranges {
        let target = range(tx, sheet_id, item)?;
        for (row, column) in coordinates(&target) {
            let mut cell = tx.read(sheet_id, row, column)?.unwrap_or_default();
            match editor {
                Some(Value::Null) | None => {
                    cell.metadata.remove("editor");
                }
                Some(value) => {
                    if !value.is_object() {
                        return Err(invalid("editor must be an object or null"));
                    }
                    cell.metadata.insert("editor".into(), value.clone());
                }
            }
            tx.write(sheet_id, row, column, Some(cell))?;
        }
    }
    Ok(true)
}

fn template_set(tx: &mut Transaction<'_>, p: &Value) -> KernelResult<bool> {
    let template = required(p, "template")?
        .as_object()
        .ok_or_else(|| invalid("cellTemplate.set template must be an object"))?;
    let id = template
        .get("id")
        .and_then(Value::as_str)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| invalid("Cell template id is required"))?;
    if template
        .get("name")
        .and_then(Value::as_str)
        .filter(|value| !value.is_empty())
        .is_none()
        || !template.get("style").map(Value::is_object).unwrap_or(false)
    {
        return Err(invalid("Cell template is invalid"));
    }
    let templates = tx
        .manifest
        .metadata
        .entry("cellStyleTemplates".into())
        .or_insert_with(|| Value::Array(Vec::new()));
    let list = templates
        .as_array_mut()
        .ok_or_else(|| invalid("cellStyleTemplates metadata is invalid"))?;
    list.retain(|value| value.get("id").and_then(Value::as_str) != Some(id));
    list.push(Value::Object(template.clone()));
    Ok(true)
}

fn template_remove(tx: &mut Transaction<'_>, p: &Value) -> KernelResult<bool> {
    let id = text(p, "templateId")?;
    let Some(value) = tx.manifest.metadata.get_mut("cellStyleTemplates") else {
        return Ok(true);
    };
    let list = value
        .as_array_mut()
        .ok_or_else(|| invalid("cellStyleTemplates metadata is invalid"))?;
    list.retain(|value| value.get("id").and_then(Value::as_str) != Some(id));
    Ok(true)
}

fn fill_applied(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let plan = fill_plan(tx, sheet_id, p)?;
    let supplied = supplied_writes(p)?;
    if !supplied.is_empty() {
        let expected: BTreeSet<_> = plan
            .iter()
            .map(|(row, column, _)| (*row, *column))
            .collect();
        let actual: BTreeSet<_> = supplied
            .iter()
            .map(|write| (write.row, write.column))
            .collect();
        if expected != actual {
            return Err(invalid(
                "Fill mutation writes do not match the canonical plan",
            ));
        }
        for write in supplied {
            let current = tx.read(sheet_id, write.row, write.column)?;
            if !cell_json_equal(current.as_ref(), write.before.as_ref()) {
                return Err(invalid(format!(
                    "Fill target changed at {}:{}",
                    write.row, write.column
                )));
            }
        }
    }
    for (row, column, cell) in plan {
        tx.write(sheet_id, row, column, cell)?;
    }
    Ok(true)
}

fn fill_restored(tx: &mut Transaction<'_>, sheet_id: &str, p: &Value) -> KernelResult<bool> {
    require_sheet_param(p, sheet_id)?;
    let writes = supplied_writes(p)?;
    if writes.is_empty() {
        return Err(invalid("fill.restored writes are required"));
    }
    for write in &writes {
        let current = tx.read(sheet_id, write.row, write.column)?;
        if !cell_json_equal(current.as_ref(), write.before.as_ref()) {
            return Err(invalid(format!(
                "Fill target changed at {}:{}",
                write.row, write.column
            )));
        }
    }
    for write in writes {
        tx.write(sheet_id, write.row, write.column, write.after)?;
    }
    Ok(true)
}

#[derive(Clone)]
struct FillWrite {
    row: u32,
    column: u32,
    before: Option<Cell>,
    after: Option<Cell>,
}

fn supplied_writes(p: &Value) -> KernelResult<Vec<FillWrite>> {
    let Some(values) = p.as_object().and_then(|object| object.get("writes")) else {
        return Ok(Vec::new());
    };
    let values = values
        .as_array()
        .ok_or_else(|| invalid("Fill writes must be an array"))?;
    values
        .iter()
        .map(|value| {
            Ok(FillWrite {
                row: integer(value, "row")?,
                column: integer(value, "column")?,
                before: parse_optional_cell(value, "before")?,
                after: parse_optional_cell(value, "after")?,
            })
        })
        .collect()
}

fn fill_plan(
    tx: &Transaction<'_>,
    sheet_id: &str,
    p: &Value,
) -> KernelResult<Vec<(u32, u32, Option<Cell>)>> {
    let source = range(tx, sheet_id, required(p, "sourceRange")?)?;
    let target = range(tx, sheet_id, required(p, "targetRange")?)?;
    let direction = text(p, "direction")?;
    let mode = text(p, "mode")?;
    if !matches!(direction, "down" | "up" | "right" | "left") || !matches!(mode, "copy" | "series")
    {
        return Err(invalid("Fill mode or direction is invalid"));
    }
    assert_fill_geometry(&source, &target, direction)?;
    if source.sheet_id != target.sheet_id {
        return Err(invalid("Fill ranges must use one worksheet"));
    }
    let mut result = Vec::new();
    let series = p.as_object().and_then(|object| object.get("series"));
    let series_type = series
        .and_then(|value| value.get("type"))
        .and_then(Value::as_str)
        .unwrap_or("linear");
    if mode == "series" && series_type != "autofill" {
        result = plan_series(tx, &source, &target, direction, series)?;
    } else {
        for (row, column) in coordinates(&target) {
            if source.contains(&CellAddress {
                sheet_id: sheet_id.into(),
                row,
                column,
            }) {
                continue;
            }
            let (source_row, source_column) = source_coordinate(&source, row, column, direction);
            let source_cell = tx.read(sheet_id, source_row, source_column)?;
            let next = copy_cell(
                source_cell,
                row as i64 - source_row as i64,
                column as i64 - source_column as i64,
            )?;
            let current = tx.read(sheet_id, row, column)?;
            if current != next {
                result.push((row, column, next));
            }
        }
    }
    Ok(result)
}

fn plan_series(
    tx: &Transaction<'_>,
    source: &RangeRef,
    target: &RangeRef,
    direction: &str,
    options: Option<&Value>,
) -> KernelResult<Vec<(u32, u32, Option<Cell>)>> {
    let step = options
        .and_then(|value| value.get("stepValue"))
        .and_then(Value::as_f64);
    let stop = options
        .and_then(|value| value.get("stopValue"))
        .and_then(Value::as_f64);
    let kind = options
        .and_then(|value| value.get("type"))
        .and_then(Value::as_str)
        .unwrap_or("linear");
    let mut tracks: HashMap<u32, Vec<(u32, u32, f64, Cell, i64, bool)>> = HashMap::new();
    for (row, column) in coordinates(source) {
        let Some(cell) = tx.read(&source.sheet_id, row, column)? else {
            continue;
        };
        if cell.formula.is_some() {
            return Err(invalid("Series fill accepts only numeric seeds"));
        }
        let date_seed = kind == "date"
            || cell
                .metadata
                .get("numberFormat")
                .and_then(Value::as_str)
                .is_some_and(is_date_format);
        let value = if date_seed {
            if let Scalar::Number(serial) = &cell.value {
                excel_to_civil_days(
                    *serial,
                    tx.manifest
                        .metadata
                        .get("date1904")
                        .and_then(Value::as_bool)
                        .unwrap_or(false),
                )
            } else {
                parse_iso_date(&cell.value)
                    .ok_or_else(|| invalid("Series fill date seed is invalid"))?
            }
        } else {
            match cell.value {
                Scalar::Number(value) if value.is_finite() => value,
                _ => return Err(invalid("Series fill accepts only finite numeric seeds")),
            }
        };
        let track = if matches!(direction, "down" | "up") {
            column
        } else {
            row
        };
        let travel = travel_coordinate(direction, row, column);
        tracks
            .entry(track)
            .or_default()
            .push((row, column, value, cell, travel, date_seed));
    }
    for seeds in tracks.values_mut() {
        seeds.sort_by_key(|seed| seed.4);
    }
    let mut result = Vec::new();
    for (row, column) in coordinates(target) {
        if source.contains(&CellAddress {
            sheet_id: source.sheet_id.clone(),
            row,
            column,
        }) {
            continue;
        }
        let track = if matches!(direction, "down" | "up") {
            column
        } else {
            row
        };
        let seeds = tracks
            .get(&track)
            .ok_or_else(|| invalid("Series fill requires a seed on every affected track"))?;
        let first = seeds
            .first()
            .ok_or_else(|| invalid("Series fill requires at least one numeric seed"))?;
        let travel = travel_coordinate(direction, row, column);
        let distance = travel - first.4;
        let value = if kind == "date" || first.5 {
            date_series_value(first.2, distance, step, options)?
        } else if kind == "growth" {
            let ratio = step.unwrap_or_else(|| {
                if seeds.len() > 1 && first.2 != 0.0 {
                    seeds[1].2 / first.2
                } else {
                    2.0
                }
            });
            first.2 * ratio.powi(distance as i32)
        } else {
            let delta = step.unwrap_or_else(|| {
                if options
                    .and_then(|value| value.get("trend"))
                    .and_then(Value::as_bool)
                    .unwrap_or(false)
                {
                    trend_delta(seeds)
                } else if seeds.len() > 1 {
                    (seeds[1].2 - first.2) / (seeds[1].4 - first.4) as f64
                } else {
                    1.0
                }
            });
            let intercept = first.2 - delta * first.4 as f64;
            intercept + delta * travel as f64
        };
        if !value.is_finite()
            || stop.is_some_and(|limit| {
                (value >= first.2 && value > limit) || (value < first.2 && value < limit)
            })
        {
            continue;
        }
        let mut next = first.3.clone();
        next.value = if kind == "date" || first.5 {
            if let Scalar::Number(original) = &first.3.value {
                let unit = options
                    .and_then(|o| o.get("dateUnit"))
                    .and_then(Value::as_str)
                    .unwrap_or("day");
                Scalar::Number(if unit == "day" {
                    *original + step.unwrap_or(1.0) * distance as f64
                } else {
                    civil_days_to_excel(
                        value,
                        tx.manifest
                            .metadata
                            .get("date1904")
                            .and_then(Value::as_bool)
                            .unwrap_or(false),
                    )
                })
            } else {
                Scalar::Text(format_iso_date(value))
            }
        } else {
            Scalar::Number(value)
        };
        next.formula = None;
        let current = tx.read(&source.sheet_id, row, column)?;
        if current != Some(next.clone()) {
            result.push((row, column, Some(next)));
        }
    }
    Ok(result)
}

fn source_coordinate(source: &RangeRef, row: u32, column: u32, direction: &str) -> (u32, u32) {
    let height = source.end_row - source.start_row + 1;
    let width = source.end_column - source.start_column + 1;
    let source_row = if direction == "up" {
        source.end_row - ((source.end_row as i64 - row as i64).rem_euclid(height as i64) as u32)
    } else {
        source.start_row + ((row - source.start_row) % height)
    };
    let source_column = if direction == "left" {
        source.end_column
            - ((source.end_column as i64 - column as i64).rem_euclid(width as i64) as u32)
    } else {
        source.start_column + ((column - source.start_column) % width)
    };
    (source_row, source_column)
}

fn travel_coordinate(direction: &str, row: u32, column: u32) -> i64 {
    let coordinate = (if matches!(direction, "down" | "up") {
        row
    } else {
        column
    }) as i64;
    if matches!(direction, "up" | "left") {
        -coordinate
    } else {
        coordinate
    }
}

fn is_date_format(format: &str) -> bool {
    let mut stripped = String::new();
    let mut quoted = false;
    for character in format.chars() {
        if character == '"' {
            quoted = !quoted;
            continue;
        }
        if !quoted && character != '\\' {
            stripped.push(character);
        }
    }
    stripped
        .chars()
        .any(|character| matches!(character, 'y' | 'Y' | 'd' | 'D' | 'h' | 'H' | 's' | 'S'))
}

fn parse_iso_date(value: &Scalar) -> Option<f64> {
    let text = match value {
        Scalar::Text(text) => text,
        Scalar::Number(value) if value.is_finite() => return Some(*value),
        _ => return None,
    };
    let bytes = text.as_bytes();
    if bytes.len() < 10 {
        return None;
    }
    let year = text.get(0..4)?.parse::<i32>().ok()?;
    let month = text.get(5..7)?.parse::<u32>().ok()?;
    let day = text.get(8..10)?.parse::<u32>().ok()?;
    if text.as_bytes().get(4) != Some(&b'-')
        || text.as_bytes().get(7) != Some(&b'-')
        || month == 0
        || month > 12
        || day == 0
        || day > days_in_month(year, month)
    {
        return None;
    }
    let mut serial = days_from_civil(year, month, day) as f64;
    if bytes.len() >= 19 {
        if let (Some(hour), Some(minute), Some(second)) = (
            text.get(11..13).and_then(|v| v.parse::<u32>().ok()),
            text.get(14..16).and_then(|v| v.parse::<u32>().ok()),
            text.get(17..19).and_then(|v| v.parse::<u32>().ok()),
        ) {
            serial += (hour * 3600 + minute * 60 + second) as f64 / 86_400.0;
        }
    }
    Some(serial)
}
fn excel_to_civil_days(serial: f64, date1904: bool) -> f64 {
    if date1904 {
        days_from_civil(1904, 1, 1) as f64 + serial
    } else {
        days_from_civil(1899, 12, 31) as f64 + serial - if serial >= 60.0 { 1.0 } else { 0.0 }
    }
}
fn civil_days_to_excel(days: f64, date1904: bool) -> f64 {
    if date1904 {
        days - days_from_civil(1904, 1, 1) as f64
    } else {
        let serial = days - days_from_civil(1899, 12, 31) as f64;
        serial + if serial >= 60.0 { 1.0 } else { 0.0 }
    }
}

fn date_series_value(
    first: f64,
    distance: i64,
    step: Option<f64>,
    options: Option<&Value>,
) -> KernelResult<f64> {
    let amount = step.unwrap_or(1.0) * distance as f64;
    if amount.fract() != 0.0 {
        return Ok(first + amount);
    }
    let unit = options
        .and_then(|value| value.get("dateUnit"))
        .and_then(Value::as_str)
        .unwrap_or("day");
    if unit == "day" {
        return Ok(first + amount);
    }
    if unit == "weekday" {
        let mut current = first;
        let direction = if amount < 0.0 { -1.0 } else { 1.0 };
        let mut remaining = amount.abs() as u64;
        while remaining > 0 {
            current += direction;
            let weekday = (days_from_epoch(current) + 4).rem_euclid(7);
            if weekday != 0 && weekday != 6 {
                remaining -= 1;
            }
        }
        return Ok(current);
    }
    let (year, month, day) = civil_from_days(first.floor() as i64);
    let month_delta = if unit == "year" {
        amount as i32 * 12
    } else {
        amount as i32
    };
    let absolute = year * 12 + month as i32 - 1 + month_delta;
    let next_year = absolute.div_euclid(12);
    let next_month = absolute.rem_euclid(12) as u32 + 1;
    let day = day.min(days_in_month(next_year, next_month));
    Ok(days_from_civil(next_year, next_month, day) as f64 + first.fract())
}

fn trend_delta(seeds: &[(u32, u32, f64, Cell, i64, bool)]) -> f64 {
    if seeds.len() < 2 {
        return 1.0;
    }
    let mean_x = seeds.iter().map(|seed| seed.4 as f64).sum::<f64>() / seeds.len() as f64;
    let mean_y = seeds.iter().map(|seed| seed.2).sum::<f64>() / seeds.len() as f64;
    let denominator = seeds
        .iter()
        .map(|seed| (seed.4 as f64 - mean_x).powi(2))
        .sum::<f64>();
    if denominator == 0.0 {
        1.0
    } else {
        seeds
            .iter()
            .map(|seed| (seed.4 as f64 - mean_x) * (seed.2 - mean_y))
            .sum::<f64>()
            / denominator
    }
}

fn format_iso_date(serial: f64) -> String {
    let (year, month, day) = civil_from_days(serial.floor() as i64);
    let milliseconds = (serial.fract().max(0.0) * 86_400_000.0).round() as u64;
    let hour = (milliseconds / 3_600_000) % 24;
    let minute = (milliseconds / 60_000) % 60;
    let second = (milliseconds / 1_000) % 60;
    let milli = milliseconds % 1_000;
    format!("{year:04}-{month:02}-{day:02}T{hour:02}:{minute:02}:{second:02}.{milli:03}Z")
}
fn days_from_epoch(serial: f64) -> i64 {
    serial.floor() as i64
}
fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}
fn days_from_civil(year: i32, month: u32, day: u32) -> i64 {
    let year = year - if month <= 2 { 1 } else { 0 };
    let era = (if year >= 0 { year } else { year - 399 }).div_euclid(400);
    let year_of_era = year - era * 400;
    let month = month as i32;
    let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + day as i32 - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era as i64 * 146097 + day_of_era as i64 - 719468
}
fn civil_from_days(days: i64) -> (i32, u32, u32) {
    let days = days + 719468;
    let era = (if days >= 0 { days } else { days - 146096 }).div_euclid(146097);
    let day_of_era = (days - era * 146097) as i32;
    let year_of_era =
        (day_of_era - day_of_era / 1460 + day_of_era / 36524 - day_of_era / 146096) / 365;
    let year = year_of_era + era as i32 * 400;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_part = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_part + 2) / 5 + 1;
    let month = month_part + if month_part < 10 { 3 } else { -9 };
    (
        year + if month <= 2 { 1 } else { 0 },
        month as u32,
        day as u32,
    )
}

fn copy_cell(cell: Option<Cell>, row_delta: i64, column_delta: i64) -> KernelResult<Option<Cell>> {
    let Some(mut cell) = cell else {
        return Ok(None);
    };
    if let Some(formula) = cell.formula.clone() {
        let row_delta = i32::try_from(row_delta)
            .map_err(|_| invalid("Formula row offset is outside supported bounds"))?;
        let column_delta = i32::try_from(column_delta)
            .map_err(|_| invalid("Formula column offset is outside supported bounds"))?;
        cell.formula = Some(kernel_formula::references::offset(
            &formula,
            row_delta,
            column_delta,
        )?);
        cell.value = Scalar::Null;
        for key in ["formulaValue", "displayValue", "formulaMetadata"] {
            cell.metadata.remove(key);
        }
    }
    Ok(Some(cell))
}

fn clear_contents(mut cell: Cell) -> Cell {
    cell.value = Scalar::Null;
    cell.formula = None;
    for key in ["formulaValue", "displayValue", "formulaMetadata"] {
        cell.metadata.remove(key);
    }
    cell
}
fn clear_formats(mut cell: Cell) -> Cell {
    for key in ["style", "styleId", "numberFormat", "displayValue"] {
        cell.metadata.remove(key);
    }
    cell
}
fn clear_metadata_family(mut cell: Cell, family: &str) -> Cell {
    if family == "comments-and-notes" {
        for key in ["comment", "comments", "note", "notes"] {
            cell.metadata.remove(key);
        }
    } else {
        for key in ["hyperlink", "hyperlinks"] {
            cell.metadata.remove(key);
        }
    }
    cell
}
fn parse_cell(value: &Value) -> KernelResult<Cell> {
    serde_json::from_value(value.clone())
        .map_err(|error| invalid(format!("Cell is invalid: {error}")))
}
fn parse_optional_cell(value: &Value, key: &str) -> KernelResult<Option<Cell>> {
    match value.as_object().and_then(|object| object.get(key)) {
        None | Some(Value::Null) => Ok(None),
        Some(value) => Ok(Some(parse_cell(value)?)),
    }
}
fn required<'a>(p: &'a Value, key: &str) -> KernelResult<&'a Value> {
    p.as_object()
        .and_then(|object| object.get(key))
        .ok_or_else(|| invalid(format!("{key} is required")))
}
fn require_sheet_param(p: &Value, sheet_id: &str) -> KernelResult<()> {
    if p.as_object()
        .and_then(|object| object.get("sheetId"))
        .and_then(Value::as_str)
        != Some(sheet_id)
    {
        return Err(invalid("sheetId must match the mutation sheet"));
    }
    Ok(())
}
fn bounds(tx: &Transaction<'_>, sheet_id: &str, row: u32, column: u32) -> KernelResult<()> {
    if row >= MAX_ROWS || column >= MAX_COLUMNS {
        return Err(invalid("Cell coordinate is outside worksheet limits"));
    }
    let sheet = tx.sheet(sheet_id)?;
    if row >= sheet.row_count || column >= sheet.column_count {
        return Err(invalid("Cell coordinate is outside worksheet bounds"));
    }
    Ok(())
}
fn offset(base: u32, offset: usize) -> KernelResult<u32> {
    base.checked_add(u32::try_from(offset).map_err(|_| invalid("Cell coordinate overflow"))?)
        .ok_or_else(|| invalid("Cell coordinate overflow"))
}
fn coordinates(range: &RangeRef) -> impl Iterator<Item = (u32, u32)> {
    (range.start_row..=range.end_row).flat_map(move |row| {
        (range.start_column..=range.end_column).map(move |column| (row, column))
    })
}
fn clear_range_cells(tx: &mut Transaction<'_>, range: &RangeRef) -> KernelResult<()> {
    for (row, column) in coordinates(range) {
        tx.write(&range.sheet_id, row, column, None)?;
    }
    Ok(())
}
fn cell_json_equal(left: Option<&Cell>, right: Option<&Cell>) -> bool {
    left == right
}
fn assert_fill_geometry(source: &RangeRef, target: &RangeRef, direction: &str) -> KernelResult<()> {
    if !target.contains(&CellAddress {
        sheet_id: source.sheet_id.clone(),
        row: source.start_row,
        column: source.start_column,
    }) || target.start_row > source.start_row
        || target.end_row < source.end_row
        || target.start_column > source.start_column
        || target.end_column < source.end_column
    {
        return Err(invalid("Fill target must contain source range"));
    }
    let same_columns =
        source.start_column == target.start_column && source.end_column == target.end_column;
    let same_rows = source.start_row == target.start_row && source.end_row == target.end_row;
    let valid = match direction {
        "down" => same_columns && target.start_row == source.start_row,
        "up" => same_columns && target.end_row == source.end_row,
        "right" => same_rows && target.start_column == source.start_column,
        "left" => same_rows && target.end_column == source.end_column,
        _ => false,
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "Fill direction requires a contiguous one-axis target extension",
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn date_series_advances_calendar_months() {
        let seed = parse_iso_date(&Scalar::Text("2026-01-31T00:00:00.000Z".into())).unwrap();
        let next = date_series_value(
            seed,
            1,
            Some(1.0),
            Some(&serde_json::json!({"dateUnit":"month"})),
        )
        .unwrap();
        assert_eq!(format_iso_date(next), "2026-02-28T00:00:00.000Z");
    }

    #[test]
    fn fill_geometry_rejects_diagonal_target() {
        let source = RangeRef {
            sheet_id: "sheet".into(),
            start_row: 0,
            end_row: 0,
            start_column: 0,
            end_column: 0,
        };
        let target = RangeRef {
            sheet_id: "sheet".into(),
            start_row: 0,
            end_row: 1,
            start_column: 0,
            end_column: 1,
        };
        assert!(assert_fill_geometry(&source, &target, "down").is_err());
    }
}
