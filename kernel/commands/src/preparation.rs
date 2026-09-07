use crate::{CommandRequest, Mutation, integer, invalid, required_role};
use kernel_core::*;
use serde_json::Value;
use std::collections::BTreeSet;

/// Resolves the immutable directory pages used by an operation before loading
/// bytes. Directory selection never probes a missing page or mistakes it for blank.
pub fn required_pages(
    pages: &WorkbookPages,
    request: &CommandRequest,
) -> KernelResult<Vec<PageKey>> {
    required_role(request)?;
    let manifest = pages.manifest();
    if request.unit_id != manifest.unit_id {
        return Err(KernelError::new(
            "WORKBOOK_ID_MISMATCH",
            "Preparation targets another workbook",
        ));
    }
    if request.base_revision != manifest.revision {
        return Err(
            KernelError::new("STALE_REVISION", "Preparation base revision is stale")
                .recover("refresh-manifest"),
        );
    }
    if request.command_id == "history.undo" {
        let _: HistoryRecord = serde_json::from_value(
            request.params.get("history").cloned().ok_or_else(|| invalid("history is required"))?,
        ).map_err(|error| invalid(format!("Invalid history record: {error}")))?;
        return Ok(Vec::new());
    }
    let mutations: Vec<Mutation> = if request.command_id == "operation.apply" {
        serde_json::from_value(request.params["mutations"].clone())
            .map_err(|e| invalid(e.to_string()))?
    } else {
        vec![Mutation {
            id: request.command_id.clone(),
            sheet_id: request.params["sheetId"].as_str().unwrap_or("").into(),
            params: request.params.clone(),
        }]
    };
    let mut keys = BTreeSet::new();
    for m in mutations {
        let p = &m.params;
        if matches!(
            m.id.as_str(),
            "rows.inserted"
                | "rows.deleted"
                | "columns.inserted"
                | "columns.deleted"
                | "cells.inserted"
                | "cells.deleted"
                | "cells.inserted.restore"
                | "cells.deleted.restore"
                | "sheet.rename"
                | "sheet.remove"
                | "sheet.duplicated"
                | "sheet.restore"
        ) {
            // Formula, defined-name, rule and object references may live on any
            // sheet. Structural commands visit every authored page once.
            keys.extend(manifest.pages.iter().map(PageDescriptor::key));
            continue;
        }
        if matches!(
            m.id.as_str(),
            "cell.set"
                | "cell.restore"
                | "note.set"
                | "note.remove"
                | "note.visibility"
                | "comment.add"
        ) {
            add_cell(
                &manifest,
                &mut keys,
                &m.sheet_id,
                integer(p, "row")?,
                integer(p, "column")?,
            )?;
        }
        if m.id == "range.set" {
            let row = integer(p, "startRow")?;
            let col = integer(p, "startColumn")?;
            let rows = p["values"]
                .as_array()
                .ok_or_else(|| invalid("range.set values must be array"))?;
            for (offset, values) in rows.iter().enumerate() {
                let values = values
                    .as_array()
                    .ok_or_else(|| invalid("range.set row must be array"))?;
                if !values.is_empty() {
                    let start = row
                        .checked_add(
                            u32::try_from(offset)
                                .map_err(|_| invalid("Range dimensions overflow"))?,
                        )
                        .ok_or_else(|| invalid("Range dimensions overflow"))?;
                    add_range(
                        &manifest,
                        &mut keys,
                        &RangeRef {
                            sheet_id: m.sheet_id.clone(),
                            start_row: start,
                            end_row: start,
                            start_column: col,
                            end_column: col
                                .checked_add(values.len() as u32 - 1)
                                .ok_or_else(|| invalid("Range dimensions overflow"))?,
                        },
                    )?;
                }
            }
        }
        if m.id == "range.paste" {
            let extent = &p["clipboard"]["sourceExtent"];
            let rows = integer(extent, "rows")?;
            let columns = integer(extent, "columns")?;
            if rows == 0 || columns == 0 {
                return Err(invalid("Clipboard dimensions must be positive"));
            }
            let transpose = p["spec"]["transpose"]
                .as_bool()
                .ok_or_else(|| invalid("Paste transpose must be Boolean"))?;
            let row = integer(&p["targetOrigin"], "row")?;
            let column = integer(&p["targetOrigin"], "column")?;
            add_range(
                &manifest,
                &mut keys,
                &RangeRef {
                    sheet_id: m.sheet_id.clone(),
                    start_row: row,
                    end_row: row
                        .checked_add(if transpose { columns - 1 } else { rows - 1 })
                        .ok_or_else(|| invalid("Paste rows overflow"))?,
                    start_column: column,
                    end_column: column
                        .checked_add(if transpose { rows - 1 } else { columns - 1 })
                        .ok_or_else(|| invalid("Paste columns overflow"))?,
                },
            )?;
        }
        collect_ranges(&manifest, &mut keys, p, &m.sheet_id)?;
        if m.id == "rows.permuted" {
            let r: RangeRef =
                serde_json::from_value(p["range"].clone()).map_err(|e| invalid(e.to_string()))?;
            let s = manifest
                .sheets
                .iter()
                .find(|s| s.sheet_id == m.sheet_id)
                .ok_or_else(|| invalid("Worksheet not found"))?;
            add_range(
                &manifest,
                &mut keys,
                &RangeRef {
                    start_column: 0,
                    end_column: s.column_count - 1,
                    ..r
                },
            )?;
            keys.extend(
                manifest
                    .pages
                    .iter()
                    .filter(|p| p.sheet_id == m.sheet_id)
                    .map(PageDescriptor::key),
            );
        }
        if m.id == "sheet.extent.restore" {
            keys.extend(
                manifest
                    .pages
                    .iter()
                    .filter(|p| p.sheet_id == m.sheet_id)
                    .map(PageDescriptor::key),
            );
        }
        // Rule formulas and list sources are canonical command inputs too.
        // Formula-backed rules can read any referenced sheet; until the formula
        // dependency service provides a request graph, select its known pages.
        if matches!(
            m.id.as_str(),
            "cell.set" | "range.set" | "range.paste" | "fill.applied"
        ) {
            if let Some(s) = manifest.sheets.iter().find(|s| s.sheet_id == m.sheet_id) {
                if let Some(rules) = s.metadata.get("dataValidations") {
                    collect_ranges(&manifest, &mut keys, rules, &s.sheet_id)?;
                    if contains_formula(rules) {
                        keys.extend(manifest.pages.iter().map(PageDescriptor::key));
                    }
                }
            }
        }
        if m.id.starts_with("query.load.") || m.id.starts_with("pivot.") {
            // Rebinding removes old source-owned regions; the definition passed
            // by the caller is not the owner of those existing ranges.
            for s in &manifest.sheets {
                for field in ["dataRegions", "pivots"] {
                    if let Some(v) = s.metadata.get(field) {
                        collect_ranges(&manifest, &mut keys, v, &s.sheet_id)?;
                    }
                }
            }
        }
        if matches!(
            m.id.as_str(),
            "comment.reply"
                | "comment.reply.remove"
                | "comment.resolve"
                | "comment.remove"
                | "comment.update"
        ) {
            let s = manifest
                .sheets
                .iter()
                .find(|s| s.sheet_id == m.sheet_id)
                .ok_or_else(|| invalid("Worksheet not found"))?;
            let t = &s
                .metadata
                .get("review")
                .ok_or_else(|| invalid("Review metadata missing"))?["threadsById"][p["threadId"]
                .as_str()
                .ok_or_else(|| invalid("threadId required"))?];
            add_cell(
                &manifest,
                &mut keys,
                &m.sheet_id,
                integer(t, "row")?,
                integer(t, "column")?,
            )?;
        }
    }
    Ok(keys.into_iter().collect())
}
fn add_cell(
    m: &WorkbookManifest,
    keys: &mut BTreeSet<PageKey>,
    sheet: &str,
    row: u32,
    column: u32,
) -> KernelResult<()> {
    let a = CellAddress {
        sheet_id: sheet.into(),
        row,
        column,
    };
    a.validate()?;
    let key = PageKey::for_address(&a);
    if m.pages.iter().any(|p| p.key() == key) {
        keys.insert(key);
    }
    Ok(())
}
fn add_range(m: &WorkbookManifest, keys: &mut BTreeSet<PageKey>, r: &RangeRef) -> KernelResult<()> {
    r.validate()?;
    for p in m.pages.iter().filter(|p| p.sheet_id == r.sheet_id) {
        if p.page_row >= r.start_row / PAGE_ROWS
            && p.page_row <= r.end_row / PAGE_ROWS
            && p.page_column >= r.start_column / PAGE_COLUMNS
            && p.page_column <= r.end_column / PAGE_COLUMNS
        {
            keys.insert(p.key());
        }
    }
    Ok(())
}
fn collect_ranges(
    m: &WorkbookManifest,
    keys: &mut BTreeSet<PageKey>,
    v: &Value,
    sheet: &str,
) -> KernelResult<()> {
    match v {
        Value::Array(values) => {
            for v in values {
                collect_ranges(m, keys, v, sheet)?;
            }
        }
        Value::Object(o) => {
            let sid = o.get("sheetId").and_then(Value::as_str).unwrap_or(sheet);
            if ["startRow", "endRow", "startColumn", "endColumn"]
                .iter()
                .all(|k| o.contains_key(*k))
            {
                add_range(
                    m,
                    keys,
                    &RangeRef {
                        sheet_id: sid.into(),
                        start_row: integer(v, "startRow")?,
                        end_row: integer(v, "endRow")?,
                        start_column: integer(v, "startColumn")?,
                        end_column: integer(v, "endColumn")?,
                    },
                )?;
            } else if o.contains_key("row") && o.contains_key("column") {
                add_cell(m, keys, sid, integer(v, "row")?, integer(v, "column")?)?;
            }
            for (key, value) in o {
                if !matches!(
                    key.as_str(),
                    "value" | "candidate" | "text" | "formula" | "payload"
                ) {
                    collect_ranges(m, keys, value, sid)?;
                }
            }
        }
        _ => {}
    }
    Ok(())
}
fn contains_formula(v: &Value) -> bool {
    match v {
        Value::Object(o) => o.iter().any(|(k, v)| {
            (matches!(k.as_str(), "formula" | "formula1" | "formula2")
                && v.as_str().is_some_and(|s| s.starts_with('=')))
                || contains_formula(v)
        }),
        Value::Array(a) => a.iter().any(contains_formula),
        _ => false,
    }
}
