use crate::Transaction;
use kernel_core::{KernelError, KernelResult, MAX_COLUMNS, MAX_ROWS, RangeRef, SheetManifest};
use serde_json::{Map, Value, json};
use std::collections::{BTreeMap, BTreeSet};

/// Canonical sheet metadata command ownership. Conditional formats and data
/// validation are handled here; conditional-format reorder remains object-owned.
pub(crate) const IDS: &[&str] = &[
    "row.hidden",
    "row.unhidden",
    "rows.unhidden.all",
    "rows.hidden.restore",
    "column.hidden",
    "column.unhidden",
    "columns.unhidden.all",
    "columns.hidden.restore",
    "autoFilter.set",
    "autoFilter.remove",
    "cf.add",
    "cf.remove",
    "cf.clear",
    "dv.add",
    "dv.remove",
    "banded.set",
    "outline.set",
    "sheetTable.add",
    "sheetTable.remove",
    "sheetTable.update",
    "sheetTable.autoFilter.set",
    "tableSheet.update",
    "ganttSheet.update",
    "reportSheet.update",
    "merge.set",
    "merge.remove",
    "freeze.set",
    "row.resize",
    "column.resize",
    "column.defaultWidth.resize",
    "columns.visibility",
    "rows.visibility",
    "view.set",
    "sheet.hidden",
    "sheet.unhidden",
    "sheet.tabColor",
    "pageLayout.margins.set",
    "pageLayout.orientation.set",
    "pageLayout.paperSize.set",
    "pageLayout.pageSetupDetail.set",
    "pageLayout.scaleToFit.set",
    "pageLayout.printTitles.set",
    "pageLayout.printArea.set",
    "pageLayout.printArea.clear",
    "pageLayout.pageBreak.insert",
    "pageLayout.pageBreak.remove",
    "pageLayout.pageBreak.clear",
    "pageLayout.printGridlines.set",
    "pageLayout.printHeadings.set",
    "pageLayout.viewGridlines.set",
    "pageLayout.viewHeadings.set",
];

pub(crate) fn supports(id: &str) -> bool {
    IDS.contains(&id)
}

pub(crate) fn apply(
    tx: &mut Transaction,
    id: &str,
    sheet_id: &str,
    p: &Value,
) -> KernelResult<bool> {
    if !supports(id) {
        return Ok(false);
    }
    let params = object(p, "Mutation params must be an object")?;
    tx.sheet(sheet_id)?;
    if id.starts_with("pageLayout.") {
        apply_page_layout(tx, id, sheet_id, params)?;
        return Ok(true);
    }
    apply_sheet(tx, id, sheet_id, params)
}

fn apply_sheet(
    tx: &mut Transaction,
    id: &str,
    sheet_id: &str,
    p: &Map<String, Value>,
) -> KernelResult<bool> {
    match id {
        "row.hidden" | "row.unhidden" => {
            let row = index(p, "row", "index")?;
            ensure_row(tx.sheet(sheet_id)?, row)?;
            set_index(
                tx.sheet_mut(sheet_id)?,
                "hiddenRows",
                row,
                id == "row.hidden",
            )?;
            affect_row(tx, sheet_id, row);
        }
        "column.hidden" | "column.unhidden" => {
            let column = index(p, "column", "index")?;
            ensure_column(tx.sheet(sheet_id)?, column)?;
            set_index(
                tx.sheet_mut(sheet_id)?,
                "hiddenColumns",
                column,
                id == "column.hidden",
            )?;
            affect_column(tx, sheet_id, column);
        }
        "rows.unhidden.all" | "rows.hidden.restore" => {
            if id == "rows.unhidden.all" {
                tx.sheet_mut(sheet_id)?
                    .metadata
                    .insert("hiddenRows".into(), Value::Array(Vec::new()));
            } else {
                let values = indices(p, "indices", tx.sheet(sheet_id)?.row_count)?;
                tx.sheet_mut(sheet_id)?
                    .metadata
                    .insert("hiddenRows".into(), json!(values));
            }
            affect_whole_sheet(tx, sheet_id)?;
        }
        "columns.unhidden.all" | "columns.hidden.restore" => {
            if id == "columns.unhidden.all" {
                tx.sheet_mut(sheet_id)?
                    .metadata
                    .insert("hiddenColumns".into(), Value::Array(Vec::new()));
            } else {
                let values = indices(p, "indices", tx.sheet(sheet_id)?.column_count)?;
                tx.sheet_mut(sheet_id)?
                    .metadata
                    .insert("hiddenColumns".into(), json!(values));
            }
            affect_whole_sheet(tx, sheet_id)?;
        }
        "columns.visibility" | "rows.visibility" => {
            let row_axis = id == "rows.visibility";
            let states = p
                .get("states")
                .and_then(Value::as_array)
                .ok_or_else(|| err("VALIDATION_ERROR", format!("{id} requires states")))?;
            if states.is_empty() {
                return Err(err("VALIDATION_ERROR", format!("{id} requires states")));
            }
            let limit = if row_axis {
                tx.sheet(sheet_id)?.row_count
            } else {
                tx.sheet(sheet_id)?.column_count
            };
            let key = if row_axis {
                "hiddenRows"
            } else {
                "hiddenColumns"
            };
            let index_key = if row_axis { "row" } else { "column" };
            let mut hidden = read_indices(&tx.sheet(sheet_id)?.metadata, key, limit)?;
            for state in states {
                let state = object(state, "Visibility state must be an object")?;
                let i = required_index(state, index_key, limit)?;
                let h = state
                    .get("hidden")
                    .and_then(Value::as_bool)
                    .ok_or_else(|| {
                        err(
                            "VALIDATION_ERROR",
                            "Visibility state hidden must be boolean",
                        )
                    })?;
                if h {
                    hidden.insert(i);
                } else {
                    hidden.remove(&i);
                }
            }
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert(key.into(), json!(hidden.into_iter().collect::<Vec<_>>()));
            for state in states {
                let i = required_index(
                    object(state, "Visibility state must be an object")?,
                    index_key,
                    limit,
                )?;
                if row_axis {
                    affect_row(tx, sheet_id, i);
                } else {
                    affect_column(tx, sheet_id, i);
                }
            }
        }
        "autoFilter.set" => {
            let filter = p
                .get("autoFilter")
                .ok_or_else(|| err("VALIDATION_ERROR", "autoFilter.set requires autoFilter"))?;
            let range = validate_filter(tx, sheet_id, filter, None)?;
            resolve_filter_owners(tx, sheet_id)?;
            ensure_filter_owner(tx, sheet_id, &range, None)?;
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert("autoFilter".into(), filter.clone());
            tx.affected.push(range);
        }
        "autoFilter.remove" => {
            resolve_filter_owners(tx, sheet_id)?;
            let old = tx.sheet(sheet_id)?.metadata.get("autoFilter").cloned();
            if let Some(value) = old.as_ref() {
                let range = crate::range(
                    tx,
                    sheet_id,
                    value
                        .get("range")
                        .ok_or_else(|| err("MANIFEST_INVALID", "AutoFilter range is missing"))?,
                )?;
                tx.affected.push(range);
            } else {
                affect_whole_sheet(tx, sheet_id)?;
            }
            tx.sheet_mut(sheet_id)?.metadata.remove("autoFilter");
        }
        "cf.add" | "dv.add" => {
            let key = if id == "cf.add" {
                "conditionalFormats"
            } else {
                "dataValidations"
            };
            let rule = p
                .get("rule")
                .ok_or_else(|| err("VALIDATION_ERROR", format!("{id} requires rule")))?;
            let ranges = rule_ranges(tx, sheet_id, rule)?;
            validate_rule(rule, id)?;
            let rule_id = required_text(object(rule, "Rule must be an object")?, "id")?;
            let mut rules = array_metadata(&tx.sheet(sheet_id)?.metadata, key)?.to_vec();
            if let Some(pos) = rules
                .iter()
                .position(|v| v.get("id").and_then(Value::as_str) == Some(rule_id))
            {
                rules[pos] = rule.clone();
            } else {
                rules.push(rule.clone());
            }
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert(key.into(), Value::Array(rules));
            tx.affected.extend(ranges);
        }
        "cf.remove" | "dv.remove" => {
            let key = if id == "cf.remove" {
                "conditionalFormats"
            } else {
                "dataValidations"
            };
            let rule_id = required_text(p, "ruleId")?;
            let rules = array_metadata(&tx.sheet(sheet_id)?.metadata, key)?.to_vec();
            let mut found = false;
            let mut next = Vec::with_capacity(rules.len());
            for rule in &rules {
                if rule.get("id").and_then(Value::as_str) == Some(rule_id) {
                    found = true;
                    let ranges = rule_ranges(tx, sheet_id, rule)?;
                    tx.affected.extend(ranges);
                } else {
                    next.push(rule.clone());
                }
            }
            if !found {
                return Err(err("NOT_FOUND", format!("Rule not found: {rule_id}")));
            }
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert(key.into(), Value::Array(next));
        }
        "cf.clear" => {
            let old = array_metadata(&tx.sheet(sheet_id)?.metadata, "conditionalFormats")?.to_vec();
            if old.is_empty() {
                affect_whole_sheet(tx, sheet_id)?;
            } else {
                for rule in &old {
                    let ranges = rule_ranges(tx, sheet_id, rule)?;
                    tx.affected.extend(ranges);
                }
            }
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert("conditionalFormats".into(), Value::Array(Vec::new()));
        }
        "banded.set" => {
            let rule = p.get("rule");
            if rule.map_or(true, Value::is_null) {
                tx.sheet_mut(sheet_id)?.metadata.remove("bandedRule");
                affect_whole_sheet(tx, sheet_id)?;
            } else {
                let rule = object(
                    rule.ok_or_else(|| err("VALIDATION_ERROR", "banded rule is required"))?,
                    "Banded rule must be an object",
                )?;
                if rule.get("sheetId").and_then(Value::as_str) != Some(sheet_id) {
                    return Err(err("VALIDATION_ERROR", "Banded rule targets another sheet"));
                }
                let range = crate::range(
                    tx,
                    sheet_id,
                    rule.get("range")
                        .ok_or_else(|| err("VALIDATION_ERROR", "Banded rule range is required"))?,
                )?;
                tx.sheet_mut(sheet_id)?
                    .metadata
                    .insert("bandedRule".into(), Value::Object(rule.clone()));
                tx.affected.push(range);
            }
        }
        "outline.set" => {
            let outline = object(
                p.get("outline").ok_or_else(|| {
                    err("VALIDATION_ERROR", "outline.set requires outline object")
                })?,
                "outline must be an object",
            )?;
            let groups = outline
                .get("groups")
                .and_then(Value::as_array)
                .ok_or_else(|| err("VALIDATION_ERROR", "outline groups must be an array"))?;
            for group in groups {
                let group = object(group, "Outline group must be an object")?;
                let _ = required_text(group, "id")?;
                let axis = required_text(group, "axis")?;
                let key = match axis {
                    "row" => "row",
                    "column" => "column",
                    _ => return Err(err("VALIDATION_ERROR", "Outline axis is invalid")),
                };
                let start = required_index(
                    group,
                    "start",
                    if key == "row" {
                        tx.sheet(sheet_id)?.row_count
                    } else {
                        tx.sheet(sheet_id)?.column_count
                    },
                )?;
                let end = required_index(
                    group,
                    "end",
                    if key == "row" {
                        tx.sheet(sheet_id)?.row_count
                    } else {
                        tx.sheet(sheet_id)?.column_count
                    },
                )?;
                if end < start
                    || group
                        .get("level")
                        .and_then(Value::as_u64)
                        .filter(|v| (1..=3).contains(v))
                        .is_none()
                    || group.get("collapsed").and_then(Value::as_bool).is_none()
                {
                    return Err(err("VALIDATION_ERROR", "Outline group is invalid"));
                }
            }
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert("outline".into(), Value::Object(outline.clone()));
            affect_whole_sheet(tx, sheet_id)?;
        }
        "sheetTable.add" | "sheetTable.update" => {
            upsert_table(tx, sheet_id, p, id == "sheetTable.add")?
        }
        "sheetTable.remove" => remove_table(tx, sheet_id, p)?,
        "sheetTable.autoFilter.set" => table_filter(tx, sheet_id, p)?,
        "tableSheet.update" | "ganttSheet.update" | "reportSheet.update" => {
            let key = match id {
                "tableSheet.update" => "tableSheet",
                "ganttSheet.update" => "ganttSheet",
                _ => "reportSheet",
            };
            let value = p
                .get("definition")
                .or_else(|| p.get("model"))
                .or_else(|| p.get("sheet"))
                .or_else(|| p.get("value"))
                .ok_or_else(|| err("VALIDATION_ERROR", format!("{id} requires model")))?;
            if !value.is_object() {
                return Err(err(
                    "VALIDATION_ERROR",
                    format!("{id} model must be an object"),
                ));
            }
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert(key.into(), value.clone());
            affect_whole_sheet(tx, sheet_id)?;
        }
        "merge.set" | "merge.remove" => merge(tx, id, sheet_id, p)?,
        "freeze.set" => freeze(tx, sheet_id, p)?,
        "row.resize" | "column.resize" => resize(tx, id, sheet_id, p)?,
        "column.defaultWidth.resize" => {
            let width = positive_number(p, "widthPx")?;
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert("defaultColumnWidthPx".into(), json!(width));
            affect_whole_sheet(tx, sheet_id)?;
        }
        "view.set" => view(tx, sheet_id, p)?,
        "sheet.hidden" | "sheet.unhidden" => {
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert("hidden".into(), Value::Bool(id == "sheet.hidden"));
            affect_whole_sheet(tx, sheet_id)?;
        }
        "sheet.tabColor" => match p.get("color") {
            None | Some(Value::Null) => {
                tx.sheet_mut(sheet_id)?.metadata.remove("tabColor");
                affect_whole_sheet(tx, sheet_id)?;
            }
            Some(v) if v.as_str().map_or(false, |s| s.len() <= 64) => {
                tx.sheet_mut(sheet_id)?
                    .metadata
                    .insert("tabColor".into(), v.clone());
                affect_whole_sheet(tx, sheet_id)?;
            }
            _ => {
                return Err(err(
                    "VALIDATION_ERROR",
                    "tabColor must be a short string or null",
                ));
            }
        },
        _ => {
            return Err(err(
                "COMMAND_UNSUPPORTED",
                format!("Unsupported sheet metadata mutation: {id}"),
            ));
        }
    }
    Ok(true)
}

fn merge(
    tx: &mut Transaction,
    id: &str,
    sheet_id: &str,
    p: &Map<String, Value>,
) -> KernelResult<()> {
    let range = crate::range(
        tx,
        sheet_id,
        p.get("range")
            .ok_or_else(|| err("VALIDATION_ERROR", "range is required"))?,
    )?;
    let current = array_metadata(&tx.sheet(sheet_id)?.metadata, "merges")?.to_vec();
    let parsed: Vec<(Value, RangeRef)> = current
        .iter()
        .map(|value| {
            let range = crate::range(
                tx,
                sheet_id,
                value
                    .get("range")
                    .ok_or_else(|| err("MANIFEST_INVALID", "Merged range is missing"))?,
            )?;
            Ok((value.clone(), range))
        })
        .collect::<KernelResult<_>>()?;
    let mut merges: Vec<Value> = parsed.iter().map(|(value, _)| value.clone()).collect();
    let anchor = (range.start_row, range.start_column);
    if id == "merge.remove" {
        let before = merges.len();
        merges = parsed
            .iter()
            .filter(|(_, r)| (r.start_row, r.start_column) != anchor)
            .map(|(value, _)| value.clone())
            .collect();
        if before == merges.len() {
            return Err(err("NOT_FOUND", "Merged range does not exist"));
        }
    } else {
        if parsed
            .iter()
            .any(|(_, existing)| existing.intersects(&range))
        {
            return Err(err("CONFLICT", "Merged ranges may not overlap"));
        }
        merges = parsed
            .iter()
            .filter(|(_, r)| (r.start_row, r.start_column) != anchor)
            .map(|(value, _)| value.clone())
            .collect();
        merges.push(json!({"range": range.clone(), "anchor": {"row": range.start_row, "column": range.start_column}}));
    }
    tx.sheet_mut(sheet_id)?
        .metadata
        .insert("merges".into(), Value::Array(merges));
    tx.affected.push(range);
    Ok(())
}

fn freeze(tx: &mut Transaction, sheet_id: &str, p: &Map<String, Value>) -> KernelResult<()> {
    let pane = p
        .get("pane")
        .ok_or_else(|| err("VALIDATION_ERROR", "freeze.set requires pane"))?;
    validate_pane(pane, tx.sheet(sheet_id)?)?;
    tx.sheet_mut(sheet_id)?
        .metadata
        .insert("pane".into(), pane.clone());
    affect_whole_sheet(tx, sheet_id)?;
    Ok(())
}

fn resize(
    tx: &mut Transaction,
    id: &str,
    sheet_id: &str,
    p: &Map<String, Value>,
) -> KernelResult<()> {
    let (index_key, value_key, collection) = if id == "row.resize" {
        ("row", "heightPx", "rowHeightsPx")
    } else {
        ("column", "widthPx", "columnWidthsPx")
    };
    let limit = if index_key == "row" {
        tx.sheet(sheet_id)?.row_count
    } else {
        tx.sheet(sheet_id)?.column_count
    };
    let i = required_index(p, index_key, limit)?;
    let value = positive_number(p, value_key)?;
    let mut map = object_metadata(&tx.sheet(sheet_id)?.metadata, collection)?;
    map.insert(i.to_string(), json!(value));
    tx.sheet_mut(sheet_id)?
        .metadata
        .insert(collection.into(), Value::Object(map));
    if index_key == "row" {
        affect_row(tx, sheet_id, i);
    } else {
        affect_column(tx, sheet_id, i);
    }
    Ok(())
}

fn view(tx: &mut Transaction, sheet_id: &str, p: &Map<String, Value>) -> KernelResult<()> {
    let mut changed = false;
    for key in ["showGridlines", "showHeaders"] {
        if let Some(value) = p.get(key) {
            if !value.is_boolean() {
                return Err(err("VALIDATION_ERROR", format!("{key} must be boolean")));
            }
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert(key.into(), value.clone());
            changed = true;
        }
    }
    if let Some(value) = p.get("zoom") {
        let zoom = value
            .as_f64()
            .filter(|v| v.is_finite() && (25.0..=400.0).contains(v))
            .ok_or_else(|| err("VALIDATION_ERROR", "zoom must be between 25 and 400"))?;
        tx.sheet_mut(sheet_id)?
            .metadata
            .insert("zoom".into(), json!(zoom));
        changed = true;
    }
    if !changed {
        return Err(err("VALIDATION_ERROR", "view.set requires a view property"));
    }
    affect_whole_sheet(tx, sheet_id)?;
    Ok(())
}

fn apply_page_layout(
    tx: &mut Transaction,
    id: &str,
    sheet_id: &str,
    p: &Map<String, Value>,
) -> KernelResult<()> {
    let mut document = print_document(tx, sheet_id)?;
    match id {
        "pageLayout.margins.set" => {
            let margins = p
                .get("margins")
                .ok_or_else(|| err("VALIDATION_ERROR", "margins is required"))?;
            validate_margins(margins)?;
            document["pageSetup"]["margins"] = margins.clone();
        }
        "pageLayout.orientation.set" => {
            let value = required_text(p, "orientation")?;
            if !matches!(value, "portrait" | "landscape") {
                return Err(err("VALIDATION_ERROR", "Page orientation is invalid"));
            }
            document["pageSetup"]["orientation"] = Value::String(value.into());
        }
        "pageLayout.paperSize.set" => {
            let value = required_text(p, "paperSize")?;
            if !matches!(value, "letter" | "a4" | "a3" | "legal" | "custom") {
                return Err(err("VALIDATION_ERROR", "Paper size is invalid"));
            }
            document["pageSetup"]["paperSize"] = Value::String(value.into());
        }
        "pageLayout.pageSetupDetail.set" => {
            let setup = p
                .get("pageSetup")
                .ok_or_else(|| err("VALIDATION_ERROR", "pageSetup is required"))?;
            validate_page_setup(setup)?;
            document["pageSetup"] = setup.clone();
        }
        "pageLayout.scaleToFit.set" => {
            let scale = positive_bounded(p, "scale", 400.0)?;
            document["pageSetup"]["scale"] = json!(scale);
            copy_or_remove(&mut document["pageSetup"], p, "fitToWidth", true)?;
            copy_or_remove(&mut document["pageSetup"], p, "fitToHeight", true)?;
        }
        "pageLayout.printTitles.set" => {
            if let Some(value) = p.get("repeatRows") {
                validate_span(value, "repeatRows")?;
            }
            if let Some(value) = p.get("repeatColumns") {
                validate_span(value, "repeatColumns")?;
            }
            copy_or_remove(&mut document, p, "repeatRows", false)?;
            copy_or_remove(&mut document, p, "repeatColumns", false)?;
        }
        "pageLayout.printArea.set" => {
            let range = crate::range(
                tx,
                sheet_id,
                p.get("range")
                    .ok_or_else(|| err("VALIDATION_ERROR", "range is required"))?,
            )?;
            document["printAreas"] = json!([{"sheetId": sheet_id, "range": range.clone()}]);
            tx.affected.push(range);
        }
        "pageLayout.printArea.clear" => {
            if let Some(areas) = p.get("printAreas") {
                validate_print_areas(tx, areas, sheet_id)?;
                document["printAreas"] = areas.clone();
            } else {
                document["printAreas"] = Value::Array(Vec::new());
            }
        }
        "pageLayout.pageBreak.insert" => {
            let br = p
                .get("pageBreak")
                .ok_or_else(|| err("VALIDATION_ERROR", "pageBreak is required"))?;
            validate_page_break(br, sheet_id)?;
            let breaks = document["pageBreaks"]
                .as_array_mut()
                .ok_or_else(|| err("MANIFEST_INVALID", "print pageBreaks must be an array"))?;
            breaks.retain(|v| v != br);
            breaks.push(br.clone());
        }
        "pageLayout.pageBreak.remove" => {
            let br = p
                .get("pageBreak")
                .ok_or_else(|| err("VALIDATION_ERROR", "pageBreak is required"))?;
            validate_page_break(br, sheet_id)?;
            if let Some(breaks) = document["pageBreaks"].as_array_mut() {
                breaks.retain(|v| v != br);
            }
        }
        "pageLayout.pageBreak.clear" => {
            if let Some(breaks) = p.get("pageBreaks") {
                validate_page_breaks(breaks, sheet_id)?;
                document["pageBreaks"] = breaks.clone();
            } else {
                document["pageBreaks"] = Value::Array(Vec::new());
            }
        }
        "pageLayout.printGridlines.set" | "pageLayout.printHeadings.set" => {
            let enabled = p
                .get("enabled")
                .and_then(Value::as_bool)
                .ok_or_else(|| err("VALIDATION_ERROR", "enabled must be boolean"))?;
            let key = if id.ends_with("printGridlines.set") {
                "printGridlines"
            } else {
                "printHeadings"
            };
            document["pageSetup"][key] = Value::Bool(enabled);
        }
        "pageLayout.viewGridlines.set" | "pageLayout.viewHeadings.set" => {
            let enabled = p
                .get("enabled")
                .and_then(Value::as_bool)
                .ok_or_else(|| err("VALIDATION_ERROR", "enabled must be boolean"))?;
            let key = if id.ends_with("viewGridlines.set") {
                "showGridlines"
            } else {
                "showHeaders"
            };
            tx.sheet_mut(sheet_id)?
                .metadata
                .insert(key.into(), Value::Bool(enabled));
        }
        _ => {
            return Err(err(
                "COMMAND_UNSUPPORTED",
                format!("Unsupported page layout mutation: {id}"),
            ));
        }
    }
    if !matches!(
        id,
        "pageLayout.viewGridlines.set" | "pageLayout.viewHeadings.set"
    ) {
        set_print_document(tx, sheet_id, document);
    }
    if !matches!(id, "pageLayout.printArea.set") {
        affect_whole_sheet(tx, sheet_id)?;
    }
    Ok(())
}

fn print_document(tx: &Transaction, sheet_id: &str) -> KernelResult<Value> {
    if let Some(Value::Array(documents)) = tx.manifest.metadata.get("printDocuments") {
        if let Some(document) = documents
            .iter()
            .find(|v| v.get("sheetId").and_then(Value::as_str) == Some(sheet_id))
        {
            return Ok(document.clone());
        }
    }
    Ok(json!({
        "schema": "PrintDocument", "unitId": tx.manifest.unit_id.clone(), "sheetId": sheet_id,
        "pageSetup": {
            "paperSize": "a4", "orientation": "portrait",
            "margins": {"top": 72, "right": 72, "bottom": 72, "left": 72, "header": 36, "footer": 36},
            "scale": 100, "printGridlines": false, "printHeadings": false,
            "centerHorizontally": false, "centerVertically": false
        },
        "printAreas": [], "pageBreaks": []
    }))
}

fn set_print_document(tx: &mut Transaction, sheet_id: &str, document: Value) {
    let docs = tx
        .manifest
        .metadata
        .entry("printDocuments".into())
        .or_insert_with(|| Value::Array(Vec::new()));
    let docs = docs
        .as_array_mut()
        .expect("printDocuments is initialized as an array");
    if let Some(pos) = docs
        .iter()
        .position(|v| v.get("sheetId").and_then(Value::as_str) == Some(sheet_id))
    {
        docs[pos] = document;
    } else {
        docs.push(document);
    }
}

fn upsert_table(
    tx: &mut Transaction,
    sheet_id: &str,
    p: &Map<String, Value>,
    adding: bool,
) -> KernelResult<()> {
    resolve_filter_owners(tx, sheet_id)?;
    let owned_table;
    let table = if let Some(table) = p.get("table") {
        table
    } else {
        owned_table = Value::Object(p.clone());
        &owned_table
    };
    let table_obj = object(table, "sheetTable must be an object")?;
    let id = required_text(table_obj, "id")?;
    let range = crate::range(
        tx,
        sheet_id,
        table_obj
            .get("range")
            .ok_or_else(|| err("VALIDATION_ERROR", "Sheet table range is missing"))?,
    )?;
    if let Some(filter) = table_obj.get("autoFilter") {
        let filter_range = validate_filter(tx, sheet_id, filter, Some(&range))?;
        ensure_filter_owner(tx, sheet_id, &filter_range, Some(id))?;
    }
    let tables = array_metadata(&tx.sheet(sheet_id)?.metadata, "sheetTables")?;
    let pos = tables
        .iter()
        .position(|v| v.get("id").and_then(Value::as_str) == Some(id));
    if adding && pos.is_some() {
        return Err(err("CONFLICT", format!("Sheet table already exists: {id}")));
    }
    if !adding && pos.is_none() {
        return Err(err("NOT_FOUND", format!("Sheet table not found: {id}")));
    }
    let mut next = tables.to_vec();
    if let Some(pos) = pos {
        next[pos] = table.clone();
    } else {
        next.push(table.clone());
    }
    tx.sheet_mut(sheet_id)?
        .metadata
        .insert("sheetTables".into(), Value::Array(next));
    tx.affected.push(range);
    Ok(())
}

fn remove_table(tx: &mut Transaction, sheet_id: &str, p: &Map<String, Value>) -> KernelResult<()> {
    let id = required_text(p, "tableId")?;
    let tables = array_metadata(&tx.sheet(sheet_id)?.metadata, "sheetTables")?;
    let old = tables
        .iter()
        .find(|v| v.get("id").and_then(Value::as_str) == Some(id))
        .cloned()
        .ok_or_else(|| err("NOT_FOUND", format!("Sheet table not found: {id}")))?;
    let mut next = Vec::with_capacity(tables.len().saturating_sub(1));
    for table in tables {
        if table.get("id").and_then(Value::as_str) != Some(id) {
            next.push(table.clone());
        }
    }
    tx.sheet_mut(sheet_id)?
        .metadata
        .insert("sheetTables".into(), Value::Array(next));
    if let Some(range) = old.get("range") {
        let range = crate::range(tx, sheet_id, range)?;
        tx.affected.push(range);
    }
    Ok(())
}

fn table_filter(tx: &mut Transaction, sheet_id: &str, p: &Map<String, Value>) -> KernelResult<()> {
    resolve_filter_owners(tx, sheet_id)?;
    let table_id = required_text(p, "tableId")?;
    let tables = array_metadata(&tx.sheet(sheet_id)?.metadata, "sheetTables")?.to_vec();
    let pos = tables
        .iter()
        .position(|v| v.get("id").and_then(Value::as_str) == Some(table_id))
        .ok_or_else(|| err("NOT_FOUND", format!("Sheet table not found: {table_id}")))?;
    let table_range = crate::range(
        tx,
        sheet_id,
        tables[pos]
            .get("range")
            .ok_or_else(|| err("MANIFEST_INVALID", "Sheet table range is missing"))?,
    )?;
    let filter = match p.get("autoFilter") {
        None | Some(Value::Null) => None,
        Some(value) => {
            let range = validate_filter(tx, sheet_id, value, Some(&table_range))?;
            ensure_filter_owner(tx, sheet_id, &range, Some(table_id))?;
            Some(value.clone())
        }
    };
    let mut table = tables[pos].clone();
    let table_obj = table
        .as_object_mut()
        .ok_or_else(|| err("MANIFEST_INVALID", "sheet table is not an object"))?;
    if let Some(filter) = filter {
        table_obj.insert("autoFilter".into(), filter);
    } else {
        table_obj.remove("autoFilter");
    }
    let mut all = array_metadata(&tx.sheet(sheet_id)?.metadata, "sheetTables")?.to_vec();
    all[pos] = table;
    tx.sheet_mut(sheet_id)?
        .metadata
        .insert("sheetTables".into(), Value::Array(all));
    tx.affected.push(table_range);
    Ok(())
}

fn validate_filter(
    tx: &Transaction,
    sheet_id: &str,
    value: &Value,
    expected: Option<&RangeRef>,
) -> KernelResult<RangeRef> {
    let filter = object(value, "AutoFilter must be an object")?;
    if filter.get("sheetId").and_then(Value::as_str) != Some(sheet_id) {
        return Err(err("VALIDATION_ERROR", "AutoFilter targets another sheet"));
    }
    let range = crate::range(
        tx,
        sheet_id,
        filter
            .get("range")
            .ok_or_else(|| err("VALIDATION_ERROR", "AutoFilter range is required"))?,
    )?;
    if let Some(expected) = expected {
        if &range != expected {
            return Err(err(
                "VALIDATION_ERROR",
                "Table AutoFilter range must equal table range",
            ));
        }
    }
    let columns = filter
        .get("columns")
        .and_then(Value::as_object)
        .ok_or_else(|| err("VALIDATION_ERROR", "AutoFilter columns must be an object"))?;
    for (key, value) in columns {
        let column = key
            .parse::<u32>()
            .map_err(|_| err("VALIDATION_ERROR", "AutoFilter column key is invalid"))?;
        let value = object(value, "AutoFilter column is invalid")?;
        if value.get("column").and_then(Value::as_u64) != Some(column as u64)
            || column < range.start_column
            || column > range.end_column
            || value.get("showButton").and_then(Value::as_bool).is_none()
            || value.get("hiddenButton").and_then(Value::as_bool).is_none()
        {
            return Err(err(
                "VALIDATION_ERROR",
                "AutoFilter column identity is invalid",
            ));
        }
        if let Some(criterion) = value.get("criterion") {
            if !criterion.is_null() {
                validate_criterion(criterion)?;
            }
        }
    }
    Ok(range)
}

fn resolve_filter_owners(tx: &Transaction, sheet_id: &str) -> KernelResult<()> {
    let sheet = tx.sheet(sheet_id)?;
    let mut owners: Vec<(Option<&str>, RangeRef)> = Vec::new();
    if let Some(filter) = sheet.metadata.get("autoFilter") {
        owners.push((
            None,
            crate::range(
                tx,
                sheet_id,
                filter.get("range").ok_or_else(|| {
                    err("MANIFEST_INVALID", "Worksheet AutoFilter range is missing")
                })?,
            )?,
        ));
    }
    if let Some(tables) = sheet.metadata.get("sheetTables") {
        let tables = tables
            .as_array()
            .ok_or_else(|| err("MANIFEST_INVALID", "sheetTables must be an array"))?;
        for table in tables {
            let table = object(table, "Sheet table must be an object")?;
            let id = required_text(table, "id")?;
            let table_range = crate::range(
                tx,
                sheet_id,
                table
                    .get("range")
                    .ok_or_else(|| err("MANIFEST_INVALID", "Sheet table range is missing"))?,
            )?;
            if let Some(filter) = table.get("autoFilter") {
                let filter_range = validate_filter(tx, sheet_id, filter, Some(&table_range))?;
                owners.push((Some(id), filter_range));
            }
        }
    }
    for left in 0..owners.len() {
        for right in (left + 1)..owners.len() {
            if owners[left].1.intersects(&owners[right].1) {
                return Err(err("VALIDATION_ERROR", "AutoFilter owners cannot overlap"));
            }
        }
    }
    Ok(())
}

fn ensure_filter_owner(
    tx: &Transaction,
    sheet_id: &str,
    candidate: &RangeRef,
    table_id: Option<&str>,
) -> KernelResult<()> {
    let sheet = tx.sheet(sheet_id)?;
    let current = match sheet.metadata.get("autoFilter") {
        None => None,
        Some(value) => Some(crate::range(
            tx,
            sheet_id,
            value
                .get("range")
                .ok_or_else(|| err("MANIFEST_INVALID", "Worksheet AutoFilter range is missing"))?,
        )?),
    };
    // A worksheet owner can replace its own filter.  A table owner cannot
    // overlap the worksheet owner; all other table owners are exclusive.
    if table_id.is_some() && current.as_ref().map_or(false, |r| r.intersects(candidate)) {
        return Err(err(
            "CONFLICT",
            "Worksheet AutoFilter overlaps an existing owner",
        ));
    }
    if let Some(Value::Array(tables)) = sheet.metadata.get("sheetTables") {
        for table in tables {
            let id = table.get("id").and_then(Value::as_str);
            if id == table_id {
                continue;
            }
            if let Some(filter) = table.get("autoFilter") {
                let range = crate::range(
                    tx,
                    sheet_id,
                    filter.get("range").ok_or_else(|| {
                        err("MANIFEST_INVALID", "Table AutoFilter range is missing")
                    })?,
                )?;
                let table_range = crate::range(
                    tx,
                    sheet_id,
                    table
                        .get("range")
                        .ok_or_else(|| err("MANIFEST_INVALID", "Sheet table range is missing"))?,
                )?;
                if range != table_range {
                    return Err(err(
                        "VALIDATION_ERROR",
                        "Table AutoFilter range must equal table range",
                    ));
                }
                if range.intersects(candidate) {
                    return Err(err("CONFLICT", "AutoFilter overlaps a table owner"));
                }
            }
        }
    }
    Ok(())
}

fn validate_criterion(value: &Value) -> KernelResult<()> {
    let criterion = object(value, "AutoFilter criterion is invalid")?;
    let kind = required_text(criterion, "kind")?;
    if !matches!(
        kind,
        "values" | "custom" | "dynamic" | "top10" | "color" | "icon"
    ) {
        return Err(err(
            "VALIDATION_ERROR",
            "AutoFilter criterion kind is invalid",
        ));
    }
    if kind == "values"
        && (!criterion.get("values").map_or(false, Value::is_array)
            || criterion
                .get("includeBlank")
                .and_then(Value::as_bool)
                .is_none())
    {
        return Err(err(
            "VALIDATION_ERROR",
            "AutoFilter values criterion is invalid",
        ));
    }
    Ok(())
}

fn validate_rule(value: &Value, id: &str) -> KernelResult<()> {
    let rule = object(value, "Rule must be an object")?;
    let _ = required_text(rule, "id")?;
    if rule.get("range").is_none() && rule.get("ranges").is_none() {
        return Err(err(
            "VALIDATION_ERROR",
            format!("{id} rule requires range or ranges"),
        ));
    }
    Ok(())
}

fn rule_ranges(tx: &Transaction, sheet_id: &str, value: &Value) -> KernelResult<Vec<RangeRef>> {
    let rule = object(value, "Rule must be an object")?;
    if let Some(range) = rule.get("range") {
        return Ok(vec![crate::range(tx, sheet_id, range)?]);
    }
    let ranges = rule
        .get("ranges")
        .and_then(Value::as_array)
        .ok_or_else(|| err("VALIDATION_ERROR", "Rule ranges must be an array"))?;
    if ranges.is_empty() {
        return Err(err("VALIDATION_ERROR", "Rule ranges must not be empty"));
    }
    ranges
        .iter()
        .map(|range| crate::range(tx, sheet_id, range))
        .collect()
}

fn validate_pane(value: &Value, sheet: &SheetManifest) -> KernelResult<()> {
    let pane = object(value, "freeze.set requires pane")?;
    let kind = required_text(pane, "kind")?;
    if !matches!(kind, "none" | "frozen" | "split") {
        return Err(err("VALIDATION_ERROR", "pane.kind is invalid"));
    }
    if kind == "none" {
        return Ok(());
    }
    let state = required_text(pane, "state")?;
    if (kind == "frozen" && !matches!(state, "frozen" | "frozenSplit"))
        || (kind == "split" && state != "split")
    {
        return Err(err("VALIDATION_ERROR", "pane.state is invalid"));
    }
    for key in ["xSplit", "ySplit"] {
        let value = pane
            .get(key)
            .and_then(Value::as_f64)
            .filter(|v| v.is_finite() && *v >= 0.0)
            .ok_or_else(|| {
                err(
                    "VALIDATION_ERROR",
                    format!("pane.{key} must be non-negative"),
                )
            })?;
        let _ = value;
    }
    let row = required_u32(pane, "startRow")?;
    let column = required_u32(pane, "startColumn")?;
    if row >= sheet.row_count || column >= sheet.column_count {
        return Err(err("VALIDATION_ERROR", "pane start is outside worksheet"));
    }
    Ok(())
}

fn validate_margins(value: &Value) -> KernelResult<()> {
    let margins = object(value, "margins must be an object")?;
    for key in ["top", "right", "bottom", "left", "header", "footer"] {
        let value = margins
            .get(key)
            .and_then(Value::as_f64)
            .filter(|v| v.is_finite() && *v >= 0.0)
            .ok_or_else(|| err("VALIDATION_ERROR", format!("margin {key} is invalid")))?;
        let _ = value;
    }
    Ok(())
}

fn validate_page_setup(value: &Value) -> KernelResult<()> {
    let setup = object(value, "pageSetup must be an object")?;
    let paper = required_text(setup, "paperSize")?;
    if !matches!(paper, "letter" | "a4" | "a3" | "legal" | "custom") {
        return Err(err("VALIDATION_ERROR", "Print paper size is invalid"));
    }
    let orientation = required_text(setup, "orientation")?;
    if !matches!(orientation, "portrait" | "landscape") {
        return Err(err("VALIDATION_ERROR", "Print orientation is invalid"));
    }
    validate_margins(
        setup
            .get("margins")
            .ok_or_else(|| err("VALIDATION_ERROR", "pageSetup margins are required"))?,
    )?;
    let _ = positive_bounded(setup, "scale", 400.0)?;
    for key in [
        "printGridlines",
        "printHeadings",
        "centerHorizontally",
        "centerVertically",
    ] {
        if setup.get(key).and_then(Value::as_bool).is_none() {
            return Err(err(
                "VALIDATION_ERROR",
                format!("pageSetup {key} must be boolean"),
            ));
        }
    }
    Ok(())
}

fn validate_print_areas(tx: &Transaction, value: &Value, sheet_id: &str) -> KernelResult<()> {
    let areas = value
        .as_array()
        .ok_or_else(|| err("VALIDATION_ERROR", "printAreas must be an array"))?;
    for area in areas {
        let area = object(area, "print area must be an object")?;
        if area.get("sheetId").and_then(Value::as_str) != Some(sheet_id) {
            return Err(err("VALIDATION_ERROR", "print area targets another sheet"));
        }
        let _ = crate::range(
            tx,
            sheet_id,
            area.get("range")
                .ok_or_else(|| err("VALIDATION_ERROR", "print area range is required"))?,
        )?;
    }
    Ok(())
}

fn validate_page_break(value: &Value, sheet_id: &str) -> KernelResult<()> {
    let value = object(value, "pageBreak must be an object")?;
    if value.get("sheetId").and_then(Value::as_str) != Some(sheet_id) {
        return Err(err("VALIDATION_ERROR", "page break targets another sheet"));
    }
    let row = value.get("row");
    let column = value.get("column");
    if row.is_some() == column.is_some() {
        return Err(err(
            "VALIDATION_ERROR",
            "page break must specify exactly one axis",
        ));
    }
    let coordinate = row
        .or(column)
        .and_then(Value::as_u64)
        .ok_or_else(|| err("VALIDATION_ERROR", "page break coordinate is invalid"))?;
    if coordinate >= MAX_ROWS as u64 {
        return Err(err(
            "VALIDATION_ERROR",
            "page break coordinate is outside worksheet",
        ));
    }
    Ok(())
}

fn validate_page_breaks(value: &Value, sheet_id: &str) -> KernelResult<()> {
    let values = value
        .as_array()
        .ok_or_else(|| err("VALIDATION_ERROR", "pageBreaks must be an array"))?;
    for value in values {
        validate_page_break(value, sheet_id)?;
    }
    Ok(())
}

fn validate_span(value: &Value, key: &str) -> KernelResult<()> {
    if value.is_null() {
        return Ok(());
    }
    let span = object(value, "Print title span must be an object")?;
    let start = required_u32(span, "start")?;
    let end = required_u32(span, "end")?;
    if end < start {
        return Err(err("VALIDATION_ERROR", format!("{key} is invalid")));
    }
    Ok(())
}

fn copy_or_remove(
    target: &mut Value,
    source: &Map<String, Value>,
    key: &str,
    positive: bool,
) -> KernelResult<()> {
    let target = target
        .as_object_mut()
        .ok_or_else(|| err("MANIFEST_INVALID", "metadata object is not an object"))?;
    if let Some(value) = source.get(key) {
        if positive && !value.is_null() && value.as_u64().filter(|v| *v > 0).is_none() {
            return Err(err("VALIDATION_ERROR", format!("{key} must be positive")));
        }
        if value.is_null() {
            target.remove(key);
        } else {
            target.insert(key.into(), value.clone());
        }
    }
    Ok(())
}

fn object<'a>(value: &'a Value, message: &str) -> KernelResult<&'a Map<String, Value>> {
    value
        .as_object()
        .ok_or_else(|| err("VALIDATION_ERROR", message))
}

fn required_text<'a>(object: &'a Map<String, Value>, key: &str) -> KernelResult<&'a str> {
    object
        .get(key)
        .and_then(Value::as_str)
        .filter(|value| !value.is_empty())
        .ok_or_else(|| err("VALIDATION_ERROR", format!("{key} is required")))
}

fn required_u32(object: &Map<String, Value>, key: &str) -> KernelResult<u32> {
    object
        .get(key)
        .and_then(Value::as_u64)
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| {
            err(
                "VALIDATION_ERROR",
                format!("{key} must be a non-negative integer"),
            )
        })
}

fn required_index(object: &Map<String, Value>, key: &str, limit: u32) -> KernelResult<u32> {
    let value = required_u32(object, key)?;
    if value >= limit {
        return Err(err("RANGE_INVALID", format!("{key} is outside worksheet")));
    }
    Ok(value)
}

fn index(object: &Map<String, Value>, primary: &str, alias: &str) -> KernelResult<u32> {
    object
        .get(primary)
        .or_else(|| object.get(alias))
        .and_then(Value::as_u64)
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| err("VALIDATION_ERROR", format!("{primary} is required")))
}

fn positive_number(object: &Map<String, Value>, key: &str) -> KernelResult<f64> {
    object
        .get(key)
        .and_then(Value::as_f64)
        .filter(|value| value.is_finite() && *value > 0.0)
        .ok_or_else(|| {
            err(
                "VALIDATION_ERROR",
                format!("{key} must be a positive number"),
            )
        })
}

fn positive_bounded(object: &Map<String, Value>, key: &str, max: f64) -> KernelResult<f64> {
    let value = positive_number(object, key)?;
    if value > max {
        return Err(err(
            "VALIDATION_ERROR",
            format!("{key} is outside allowed range"),
        ));
    }
    Ok(value)
}

fn indices(object: &Map<String, Value>, key: &str, limit: u32) -> KernelResult<Vec<u32>> {
    let values = object
        .get(key)
        .and_then(Value::as_array)
        .ok_or_else(|| err("VALIDATION_ERROR", format!("{key} must be an array")))?;
    let mut set = BTreeSet::new();
    for value in values {
        let index = value
            .as_u64()
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                err(
                    "VALIDATION_ERROR",
                    format!("{key} contains an invalid index"),
                )
            })?;
        if index >= limit {
            return Err(err(
                "RANGE_INVALID",
                format!("index {index} is outside worksheet"),
            ));
        }
        set.insert(index);
    }
    Ok(set.into_iter().collect())
}

fn read_indices(
    metadata: &BTreeMap<String, Value>,
    key: &str,
    limit: u32,
) -> KernelResult<BTreeSet<u32>> {
    let Some(value) = metadata.get(key) else {
        return Ok(BTreeSet::new());
    };
    let values = value
        .as_array()
        .ok_or_else(|| err("MANIFEST_INVALID", format!("{key} must be an array")))?;
    let mut set = BTreeSet::new();
    for value in values {
        let index = value
            .as_u64()
            .and_then(|value| u32::try_from(value).ok())
            .ok_or_else(|| {
                err(
                    "MANIFEST_INVALID",
                    format!("{key} contains an invalid index"),
                )
            })?;
        if index >= limit {
            return Err(err(
                "MANIFEST_INVALID",
                format!("{key} is outside worksheet"),
            ));
        }
        set.insert(index);
    }
    Ok(set)
}

fn array_metadata<'a>(
    metadata: &'a BTreeMap<String, Value>,
    key: &str,
) -> KernelResult<&'a Vec<Value>> {
    match metadata.get(key) {
        None => {
            static EMPTY: Vec<Value> = Vec::new();
            Ok(&EMPTY)
        }
        Some(Value::Array(values)) => Ok(values),
        Some(_) => Err(err("MANIFEST_INVALID", format!("{key} must be an array"))),
    }
}

fn object_metadata(
    metadata: &BTreeMap<String, Value>,
    key: &str,
) -> KernelResult<Map<String, Value>> {
    match metadata.get(key) {
        None => Ok(Map::new()),
        Some(Value::Object(value)) => Ok(value.clone()),
        Some(_) => Err(err("MANIFEST_INVALID", format!("{key} must be an object"))),
    }
}

fn set_index(sheet: &mut SheetManifest, key: &str, index: u32, hidden: bool) -> KernelResult<()> {
    let limit = if key == "hiddenRows" {
        sheet.row_count
    } else {
        sheet.column_count
    };
    let mut values = read_indices(&sheet.metadata, key, limit)?;
    if hidden {
        values.insert(index);
    } else {
        values.remove(&index);
    }
    sheet
        .metadata
        .insert(key.into(), json!(values.into_iter().collect::<Vec<_>>()));
    Ok(())
}

fn ensure_row(sheet: &SheetManifest, row: u32) -> KernelResult<()> {
    if row >= sheet.row_count {
        Err(err("RANGE_INVALID", "row is outside worksheet"))
    } else {
        Ok(())
    }
}

fn ensure_column(sheet: &SheetManifest, column: u32) -> KernelResult<()> {
    if column >= sheet.column_count {
        Err(err("RANGE_INVALID", "column is outside worksheet"))
    } else {
        Ok(())
    }
}

fn affect_row(tx: &mut Transaction, sheet_id: &str, row: u32) {
    if let Some(end_column) = tx
        .sheet(sheet_id)
        .ok()
        .map(|sheet| sheet.column_count.saturating_sub(1).min(MAX_COLUMNS - 1))
    {
        tx.affected.push(RangeRef {
            sheet_id: sheet_id.into(),
            start_row: row,
            end_row: row,
            start_column: 0,
            end_column,
        });
    }
}

fn affect_column(tx: &mut Transaction, sheet_id: &str, column: u32) {
    if let Some(end_row) = tx
        .sheet(sheet_id)
        .ok()
        .map(|sheet| sheet.row_count.saturating_sub(1).min(MAX_ROWS - 1))
    {
        tx.affected.push(RangeRef {
            sheet_id: sheet_id.into(),
            start_row: 0,
            end_row,
            start_column: column,
            end_column: column,
        });
    }
}

fn affect_whole_sheet(tx: &mut Transaction, sheet_id: &str) -> KernelResult<()> {
    let (row_end, column_end) = {
        let sheet = tx.sheet(sheet_id)?;
        (
            sheet.row_count.saturating_sub(1),
            sheet.column_count.saturating_sub(1),
        )
    };
    tx.affected.push(RangeRef {
        sheet_id: sheet_id.into(),
        start_row: 0,
        end_row: row_end,
        start_column: 0,
        end_column: column_end,
    });
    Ok(())
}

fn err(code: &str, message: impl Into<String>) -> KernelError {
    KernelError::new(code, message).recover("correct-input")
}
