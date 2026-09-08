//! Reducers for workbook and worksheet owned objects.
//!
//! Object state is kept in the manifest metadata maps.  The reducers in this
//! module intentionally enumerate their command ids; callers cannot use this
//! module as an arbitrary JSON patch endpoint.

use crate::Transaction;
use kernel_core::{Cell, KernelError, KernelResult, RangeRef, Scalar};
use serde_json::{Map, Value, json};
use std::collections::{BTreeMap, BTreeSet};

const DRAWING_IDS: &[&str] = &[
    "drawing.add",
    "drawing.remove",
    "drawing.transform",
    "drawing.transform.batch",
    "drawing.anchor",
    "drawing.payload.update",
    "drawing.zorder",
    "drawing.zorder.restore",
    "drawing.visibility.set",
    "drawing.rename",
];
const SPARKLINE_IDS: &[&str] = &[
    "sparkline.add",
    "sparkline.remove",
    "sparkline.update",
    "sparkline.group.add",
    "sparkline.group.remove",
    "sparkline.group.replace",
];
const DATASOURCE_IDS: &[&str] = &[
    "dataSource.add",
    "dataSource.update",
    "dataSource.remove",
    "dataRegion.add",
    "dataRegion.remove",
];
const PIVOT_IDS: &[&str] = &[
    "pivot.add",
    "pivot.remove",
    "pivot.update",
    "pivot.refresh",
    "pivot.drilldown.add",
    "pivot.drilldown.remove",
];
const QUERY_IDS: &[&str] = &[
    "query.definition.replace",
    "query.load.range",
    "query.load.sheet-table",
    "query.load.pivot-source",
    "query.load.workbook-table",
];
const STATE_IDS: &[&str] = &[
    "table.add",
    "table.remove",
    "name.set",
    "name.remove",
    "workbook.calculation.mode.set",
];

/// Apply one of the object mutations.  `false` means that the id belongs to a
/// different command family and lets the command dispatcher try its owner.
pub(crate) fn apply(
    tx: &mut Transaction,
    id: &str,
    sheet_id: &str,
    p: &Value,
) -> KernelResult<bool> {
    if id == "cf.reorder" {
        return cf_reorder(tx, sheet_id, p).map(|_| true);
    }
    if DRAWING_IDS.contains(&id) {
        drawing(tx, id, sheet_id, p)?;
        return Ok(true);
    }
    if SPARKLINE_IDS.contains(&id) {
        sparkline(tx, id, sheet_id, p)?;
        return Ok(true);
    }
    if DATASOURCE_IDS.contains(&id) {
        data_source(tx, id, sheet_id, p)?;
        return Ok(true);
    }
    if PIVOT_IDS.contains(&id) {
        pivot(tx, id, sheet_id, p)?;
        return Ok(true);
    }
    if QUERY_IDS.contains(&id) {
        query(tx, id, sheet_id, p)?;
        return Ok(true);
    }
    if STATE_IDS.contains(&id) {
        workbook_state(tx, id, sheet_id, p)?;
        return Ok(true);
    }
    Ok(false)
}

fn error(code: &str, message: impl Into<String>) -> KernelError {
    KernelError::new(code, message)
}
fn validation(message: impl Into<String>) -> KernelError {
    error("VALIDATION_ERROR", message)
}
fn required_object<'a>(p: &'a Value, key: &str) -> KernelResult<&'a Map<String, Value>> {
    p.get(key)
        .and_then(Value::as_object)
        .ok_or_else(|| validation(format!("{key} must be an object")))
}
fn required_array<'a>(p: &'a Value, key: &str) -> KernelResult<&'a Vec<Value>> {
    p.get(key)
        .and_then(Value::as_array)
        .ok_or_else(|| validation(format!("{key} must be an array")))
}
fn required_string(p: &Value, key: &str) -> KernelResult<String> {
    let value = p
        .get(key)
        .and_then(Value::as_str)
        .filter(|v| !v.is_empty())
        .ok_or_else(|| validation(format!("{key} is required")))?;
    Ok(value.to_owned())
}
fn required_bool(p: &Value, key: &str) -> KernelResult<bool> {
    p.get(key)
        .and_then(Value::as_bool)
        .ok_or_else(|| validation(format!("{key} must be boolean")))
}
fn required_u32(p: &Value, key: &str) -> KernelResult<u32> {
    let n = p
        .get(key)
        .and_then(Value::as_u64)
        .ok_or_else(|| validation(format!("{key} must be a non-negative integer")))?;
    u32::try_from(n).map_err(|_| validation(format!("{key} is too large")))
}
fn required_f64(p: &Value, key: &str) -> KernelResult<f64> {
    let n = p
        .get(key)
        .and_then(Value::as_f64)
        .ok_or_else(|| validation(format!("{key} must be a finite number")))?;
    if !n.is_finite() {
        return Err(validation(format!("{key} must be a finite number")));
    }
    Ok(n)
}
fn id_of(v: &Value) -> KernelResult<&str> {
    v.get("id")
        .and_then(Value::as_str)
        .filter(|s| !s.is_empty())
        .ok_or_else(|| validation("object id is required"))
}
fn same_sheet(v: &Value, sheet_id: &str) -> KernelResult<()> {
    if let Some(id) = v.get("sheetId").and_then(Value::as_str) {
        if id != sheet_id {
            return Err(validation("object sheetId does not match command sheet"));
        }
    }
    Ok(())
}
fn ensure_array<'a>(
    meta: &'a mut std::collections::BTreeMap<String, Value>,
    key: &str,
) -> KernelResult<&'a mut Vec<Value>> {
    let v = meta
        .entry(key.to_owned())
        .or_insert_with(|| Value::Array(Vec::new()));
    v.as_array_mut()
        .ok_or_else(|| validation(format!("metadata.{key} must be an array")))
}
fn ensure_object<'a>(
    meta: &'a mut std::collections::BTreeMap<String, Value>,
    key: &str,
) -> KernelResult<&'a mut Map<String, Value>> {
    let v = meta
        .entry(key.to_owned())
        .or_insert_with(|| Value::Object(Map::new()));
    v.as_object_mut()
        .ok_or_else(|| validation(format!("metadata.{key} must be an object")))
}
fn find_index(values: &[Value], id: &str) -> Option<usize> {
    values
        .iter()
        .position(|v| v.get("id").and_then(Value::as_str) == Some(id))
}
fn remove_id(values: &mut Vec<Value>, id: &str, label: &str) -> KernelResult<Value> {
    let i = find_index(values, id)
        .ok_or_else(|| error("NOT_FOUND", format!("{label} not found: {id}")))?;
    Ok(values.remove(i))
}
fn object_clone(v: &Value) -> KernelResult<Value> {
    if v.is_object() {
        Ok(v.clone())
    } else {
        Err(validation("object value is required"))
    }
}
fn range_from(v: &Value, sheet_id: &str) -> KernelResult<RangeRef> {
    let o = v
        .as_object()
        .ok_or_else(|| validation("range must be an object"))?;
    let sheet = o
        .get("sheetId")
        .and_then(Value::as_str)
        .unwrap_or(sheet_id)
        .to_owned();
    if sheet != sheet_id {
        return Err(validation("range targets another sheet"));
    }
    let r = RangeRef {
        sheet_id: sheet,
        start_row: number_u32(o, "startRow")?,
        end_row: number_u32(o, "endRow")?,
        start_column: number_u32(o, "startColumn")?,
        end_column: number_u32(o, "endColumn")?,
    };
    r.validate()?;
    Ok(r)
}
fn range_any(v: &Value) -> KernelResult<RangeRef> {
    let o = v
        .as_object()
        .ok_or_else(|| validation("range must be an object"))?;
    let sheet_id = o
        .get("sheetId")
        .and_then(Value::as_str)
        .filter(|s| !s.is_empty())
        .ok_or_else(|| validation("range.sheetId is required"))?
        .to_owned();
    let r = RangeRef {
        sheet_id,
        start_row: number_u32(o, "startRow")?,
        end_row: number_u32(o, "endRow")?,
        start_column: number_u32(o, "startColumn")?,
        end_column: number_u32(o, "endColumn")?,
    };
    r.validate()?;
    Ok(r)
}
fn number_u32(o: &Map<String, Value>, key: &str) -> KernelResult<u32> {
    let n = o
        .get(key)
        .and_then(Value::as_u64)
        .ok_or_else(|| validation(format!("range.{key} must be a non-negative integer")))?;
    u32::try_from(n).map_err(|_| validation(format!("range.{key} is too large")))
}
fn affected(tx: &mut Transaction, r: RangeRef) {
    tx.affected.push(r);
}
fn whole_sheet(tx: &Transaction, sheet_id: &str) -> KernelResult<RangeRef> {
    let s = tx.sheet(sheet_id)?;
    Ok(RangeRef {
        sheet_id: sheet_id.to_owned(),
        start_row: 0,
        end_row: s.row_count - 1,
        start_column: 0,
        end_column: s.column_count - 1,
    })
}

fn cf_reorder(tx: &mut Transaction, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let ids = required_array(p, "ruleIds")?;
    let sheet = tx.sheet_mut(sheet_id)?;
    let rules = ensure_array(&mut sheet.metadata, "conditionalFormats")?;
    let mut seen = BTreeSet::new();
    let mut next = Vec::with_capacity(rules.len());
    for id in ids {
        let id = id
            .as_str()
            .ok_or_else(|| validation("cf.reorder ruleIds must be strings"))?;
        if !seen.insert(id.to_owned()) {
            return Err(validation("cf.reorder ruleIds must be unique"));
        }
        let i = find_index(rules, id).ok_or_else(|| {
            error(
                "NOT_FOUND",
                format!("Conditional format rule not found: {id}"),
            )
        })?;
        next.push(rules[i].clone());
    }
    for r in rules.iter() {
        if let Some(id) = r.get("id").and_then(Value::as_str) {
            if seen.insert(id.to_owned()) {
                next.push(r.clone());
            }
        }
    }
    for (i, r) in next.iter_mut().enumerate() {
        if let Some(o) = r.as_object_mut() {
            o.insert("priority".into(), json!(i + 1));
        }
    }
    *rules = next;
    let r = whole_sheet(tx, sheet_id)?;
    affected(tx, r);
    Ok(())
}

fn drawing(tx: &mut Transaction, id: &str, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let sheet = tx.sheet_mut(sheet_id)?;
    let mut drawings_value = sheet.metadata.remove("drawings").unwrap_or_else(|| json!([]));
    let mut payloads_value = sheet.metadata.remove("drawingPayloads").unwrap_or_else(|| json!({}));
    let drawings = drawings_value.as_array_mut().ok_or_else(|| validation("drawings must be an array"))?;
    let payloads = payloads_value.as_object_mut().ok_or_else(|| validation("drawingPayloads must be an object"))?;
    match id {
        "drawing.add" => {
            let drawing = object_clone(
                p.get("drawing")
                    .ok_or_else(|| validation("drawing is required"))?,
            )?;
            let payload = object_clone(
                p.get("payload")
                    .ok_or_else(|| validation("payload is required"))?,
            )?;
            same_sheet(&drawing, sheet_id)?;
            same_sheet(&payload, sheet_id)?;
            let did = id_of(&drawing)?.to_owned();
            let pid = drawing
                .get("payloadId")
                .and_then(Value::as_str)
                .ok_or_else(|| validation("drawing.payloadId is required"))?
                .to_owned();
            if find_index(drawings, &did).is_some() {
                return Err(error("CONFLICT", format!("Drawing already exists: {did}")));
            }
            if payloads.contains_key(&pid) {
                return Err(error(
                    "CONFLICT",
                    format!("Drawing payload already exists: {pid}"),
                ));
            }
            if payload.get("kind") != drawing.get("kind") {
                return Err(validation("drawing and payload kinds must match"));
            }
            drawings.push(drawing);
            payloads.insert(pid, payload);
        }
        "drawing.remove" => {
            let did = required_string(p, "drawingId")?;
            let d = remove_id(drawings, &did, "Drawing")?;
            let pid = d
                .get("payloadId")
                .and_then(Value::as_str)
                .ok_or_else(|| validation("drawing.payloadId is required"))?;
            if payloads.remove(pid).is_none() {
                return Err(validation(format!("Drawing payload is missing: {pid}")));
            }
        }
        "drawing.transform" => {
            let did = required_string(p, "drawingId")?;
            let transform = object_clone(
                p.get("transform")
                    .ok_or_else(|| validation("transform is required"))?,
            )?;
            let d = drawing_mut(drawings, &did)?;
            d.as_object_mut()
                .unwrap()
                .insert("transform".into(), transform);
        }
        "drawing.transform.batch" => {
            let entries = required_array(p, "entries")?;
            if entries.is_empty() || entries.len() > 1000 {
                return Err(validation("drawing transform batch is invalid"));
            }
            for e in entries {
                let o = e
                    .as_object()
                    .ok_or_else(|| validation("drawing transform entry must be an object"))?;
                let did = required_string(e, "drawingId")?;
                let before = o
                    .get("before")
                    .ok_or_else(|| validation("drawing transform before is required"))?;
                let d = drawing_mut(drawings, &did)?;
                if d.get("transform") != Some(before) {
                    return Err(error(
                        "CONFLICT",
                        format!("Drawing changed before transform batch: {did}"),
                    ));
                }
            }
            for e in entries {
                let o = e.as_object().unwrap();
                let did = required_string(e, "drawingId")?;
                let after = object_clone(
                    o.get("after")
                        .ok_or_else(|| validation("drawing transform after is required"))?,
                )?;
                drawing_mut(drawings, &did)?
                    .as_object_mut()
                    .unwrap()
                    .insert("transform".into(), after);
            }
        }
        "drawing.anchor" => {
            let did = required_string(p, "drawingId")?;
            let anchor = object_clone(
                p.get("anchor")
                    .ok_or_else(|| validation("anchor is required"))?,
            )?;
            drawing_mut(drawings, &did)?
                .as_object_mut()
                .unwrap()
                .insert("anchor".into(), anchor);
        }
        "drawing.payload.update" => {
            let pid = required_string(p, "payloadId")?;
            let before = p
                .get("before")
                .ok_or_else(|| validation("before is required"))?;
            let after = object_clone(
                p.get("after")
                    .ok_or_else(|| validation("after is required"))?,
            )?;
            if payloads.get(&pid) != Some(before) {
                return Err(error(
                    "CONFLICT",
                    format!("Drawing payload changed before update: {pid}"),
                ));
            }
            if !payloads.contains_key(&pid) {
                return Err(error(
                    "NOT_FOUND",
                    format!("Drawing payload not found: {pid}"),
                ));
            }
            payloads.insert(pid, after);
        }
        "drawing.zorder" => {
            let did = required_string(p, "drawingId")?;
            let direction = required_string(p, "direction")?;
            if !matches!(
                direction.as_str(),
                "forward" | "backward" | "front" | "back"
            ) {
                return Err(validation("Drawing z-order direction is invalid"));
            }
            let i = find_index(drawings, &did)
                .ok_or_else(|| error("NOT_FOUND", format!("Drawing not found: {did}")))?;
            if direction == "front" || direction == "back" {
                let z = drawings
                    .iter()
                    .filter_map(|v| v.get("zIndex").and_then(Value::as_f64));
                let z = if direction == "front" {
                    z.fold(f64::NEG_INFINITY, f64::max) + 1.0
                } else {
                    z.fold(f64::INFINITY, f64::min) - 1.0
                };
                drawings[i]
                    .as_object_mut()
                    .unwrap()
                    .insert("zIndex".into(), json!(z));
            } else {
                let j = if direction == "forward" {
                    i + 1
                } else {
                    i.wrapping_sub(1)
                };
                if j < drawings.len() {
                    let a = drawings[i].get("zIndex").cloned();
                    let b = drawings[j].get("zIndex").cloned();
                    drawings[i]
                        .as_object_mut()
                        .unwrap()
                        .insert("zIndex".into(), b.unwrap_or(Value::Null));
                    drawings[j]
                        .as_object_mut()
                        .unwrap()
                        .insert("zIndex".into(), a.unwrap_or(Value::Null));
                }
            }
        }
        "drawing.zorder.restore" => {
            for e in required_array(p, "entries")? {
                let did = required_string(e, "drawingId")?;
                let z = required_f64(e, "zIndex")?;
                drawing_mut(drawings, &did)?
                    .as_object_mut()
                    .unwrap()
                    .insert("zIndex".into(), json!(z));
            }
        }
        "drawing.visibility.set" => {
            let did = required_string(p, "drawingId")?;
            let v = required_bool(p, "visible")?;
            drawing_mut(drawings, &did)?
                .as_object_mut()
                .unwrap()
                .insert("visible".into(), json!(v));
        }
        "drawing.rename" => {
            let did = required_string(p, "drawingId")?;
            let d = drawing_mut(drawings, &did)?.as_object_mut().unwrap();
            match p
                .get("name")
                .and_then(Value::as_str)
                .filter(|s| !s.trim().is_empty())
            {
                Some(v) => {
                    d.insert("name".into(), json!(v));
                }
                None => {
                    d.remove("name");
                }
            }
        }
        _ => unreachable!(),
    }
    sheet.metadata.insert("drawings".into(), drawings_value);
    sheet.metadata.insert("drawingPayloads".into(), payloads_value);
    let r = whole_sheet(tx, sheet_id)?;
    affected(tx, r);
    Ok(())
}
fn drawing_mut<'a>(drawings: &'a mut Vec<Value>, id: &str) -> KernelResult<&'a mut Value> {
    let i = find_index(drawings, id)
        .ok_or_else(|| error("NOT_FOUND", format!("Drawing not found: {id}")))?;
    Ok(&mut drawings[i])
}

fn sparkline(tx: &mut Transaction, id: &str, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let sheet = tx.sheet_mut(sheet_id)?;
    let mut sparklines_value = sheet.metadata.remove("sparklines").unwrap_or_else(|| json!([]));
    let mut groups_value = sheet.metadata.remove("sparklineGroups").unwrap_or_else(|| json!([]));
    let sparklines = sparklines_value.as_array_mut().ok_or_else(|| validation("sparklines must be an array"))?;
    let groups = groups_value.as_array_mut().ok_or_else(|| validation("sparklineGroups must be an array"))?;
    match id {
        "sparkline.add" => {
            let mut value = object_clone(
                p.get("sparkline")
                    .ok_or_else(|| validation("sparkline is required"))?,
            )?;
            same_sheet(&value, sheet_id)?;
            let sid = id_of(&value)?.to_owned();
            if find_index(sparklines, &sid).is_some() {
                return Err(error(
                    "CONFLICT",
                    format!("Sparkline already exists: {sid}"),
                ));
            }
            if p.get("groupState").is_some_and(|v| !v.is_null()) {
                if let Some(o) = value.as_object_mut() {
                    o.remove("groupId");
                    o.remove("showAxis");
                    o.remove("showMarkers");
                }
                apply_group_state(groups, sparklines, sheet_id, p.get("groupState").unwrap())?;
            } else if value.get("groupId").is_some() {
                return Err(validation("Sparkline group membership requires groupState"));
            }
            sparklines.push(value);
        }
        "sparkline.remove" => {
            let sid = required_string(p, "sparklineId")?;
            let i = find_index(sparklines, &sid)
                .ok_or_else(|| error("NOT_FOUND", format!("Sparkline not found: {sid}")))?;
            if sparklines[i].get("groupId").is_some()
                && p.get("groupState").map_or(true, Value::is_null)
            {
                return Err(validation("Removing grouped sparkline requires groupState"));
            }
            if let Some(state) = p.get("groupState").filter(|v| !v.is_null()) {
                apply_group_state(groups, sparklines, sheet_id, state)?;
            }
            sparklines.remove(i);
        }
        "sparkline.update" => {
            let sid = required_string(p, "sparklineId")?;
            let patch = required_object(p, "patch")?;
            let i = find_index(sparklines, &sid)
                .ok_or_else(|| error("NOT_FOUND", format!("Sparkline not found: {sid}")))?;
            let o = sparklines[i]
                .as_object_mut()
                .ok_or_else(|| validation("Sparkline must be an object"))?;
            for (k, v) in patch {
                o.insert(k.clone(), v.clone());
            }
            o.insert("sheetId".into(), json!(sheet_id));
        }
        "sparkline.group.add" | "sparkline.group.remove" | "sparkline.group.replace" => {
            apply_group_state(groups, sparklines, sheet_id, p)?
        }
        _ => unreachable!(),
    }
    sheet.metadata.insert("sparklines".into(), sparklines_value);
    sheet.metadata.insert("sparklineGroups".into(), groups_value);
    let r = whole_sheet(tx, sheet_id)?;
    affected(tx, r);
    Ok(())
}
fn apply_group_state(
    groups: &mut Vec<Value>,
    sparklines: &mut Vec<Value>,
    sheet_id: &str,
    p: &Value,
) -> KernelResult<()> {
    same_sheet(p, sheet_id)?;
    let ids = required_array(p, "groupIds")?;
    let entries = required_array(p, "groups")?;
    let members = required_array(p, "members")?;
    let allowed: BTreeSet<String> = ids
        .iter()
        .map(|v| {
            v.as_str()
                .map(str::to_owned)
                .ok_or_else(|| validation("Sparkline groupIds are invalid"))
        })
        .collect::<KernelResult<_>>()?;
    let mut next = groups
        .iter()
        .filter(|g| !allowed.contains(g.get("id").and_then(Value::as_str).unwrap_or_default()))
        .cloned()
        .collect::<Vec<_>>();
    let mut inserts = Vec::new();
    for e in entries {
        let o = e
            .as_object()
            .ok_or_else(|| validation("Sparkline group entry must be an object"))?;
        let g = object_clone(
            o.get("group")
                .ok_or_else(|| validation("Sparkline group is required"))?,
        )?;
        if !allowed.contains(id_of(&g)?) {
            return Err(validation("Sparkline group is missing from groupIds"));
        }
        inserts.push((
            o.get("index").and_then(Value::as_u64).unwrap_or(0) as usize,
            g,
        ));
    }
    inserts.sort_by_key(|x| x.0);
    for (index, g) in inserts {
        next.insert(index.min(next.len()), g);
    }
    *groups = next;
    for m in members {
        let sid = required_string(m, "sparklineId")?;
        let i = find_index(sparklines, &sid)
            .ok_or_else(|| error("NOT_FOUND", format!("Sparkline not found: {sid}")))?;
        let o = sparklines[i].as_object_mut().unwrap();
        for k in [
            "type",
            "groupId",
            "showAxis",
            "showMarkers",
            "lineWeight",
            "dateAxis",
            "dataOrientation",
            "rightToLeft",
            "hiddenCells",
            "emptyCells",
            "verticalAxis",
            "axisColor",
            "firstColor",
            "lastColor",
            "highColor",
            "lowColor",
            "negativeColor",
            "markerColor",
        ] {
            if let Some(v) = m.get(k) {
                if v.is_null() {
                    o.remove(k);
                } else {
                    o.insert(k.into(), v.clone());
                }
            }
        }
    }
    Ok(())
}

fn data_source(tx: &mut Transaction, id: &str, sheet_id: &str, p: &Value) -> KernelResult<()> {
    let mut ranges = Vec::new();
    if id.starts_with("dataSource.") {
        match id {
            "dataSource.add" | "dataSource.update" => {
                let source = object_clone(
                    p.get("source")
                        .ok_or_else(|| validation("source is required"))?,
                )?;
                let sid = id_of(&source)?.to_owned();
                same_sheet(&source, sheet_id)?;
                if let Some(source_sheet) = source.get("sourceSheetId").and_then(Value::as_str) {
                    if source_sheet != sheet_id {
                        return Err(validation(
                            "Data source sourceSheetId does not match mutation sheet",
                        ));
                    }
                }
                source_range(tx, &source, &mut ranges)?;
                let sources = ensure_array(&mut tx.manifest.metadata, "sources")?;
                if id == "dataSource.add" {
                    if find_index(sources, &sid).is_some() {
                        return Err(error(
                            "CONFLICT",
                            format!("Data source already exists: {sid}"),
                        ));
                    }
                    sources.push(source);
                } else {
                    let i = find_index(sources, &sid).ok_or_else(|| {
                        error("NOT_FOUND", format!("Data source not found: {sid}"))
                    })?;
                    sources[i] = source;
                }
            }
            "dataSource.remove" => {
                let sid = required_string(p, "sourceId")?;
                let existing = tx
                    .manifest
                    .metadata
                    .get("sources")
                    .and_then(Value::as_array)
                    .and_then(|v| find_index(v, &sid).map(|i| v[i].clone()))
                    .ok_or_else(|| error("NOT_FOUND", format!("Data source not found: {sid}")))?;
                source_range(tx, &existing, &mut ranges)?;
                let referenced = tx.manifest.sheets.iter().any(|s| {
                    s.metadata
                        .get("dataRegions")
                        .and_then(Value::as_array)
                        .is_some_and(|rs| {
                            rs.iter().any(|r| {
                                r.get("sourceId").and_then(Value::as_str) == Some(sid.as_str())
                            })
                        })
                });
                if referenced {
                    return Err(error(
                        "CONFLICT",
                        format!("Data source is still referenced by a sheet region: {sid}"),
                    ));
                }
                remove_id(
                    ensure_array(&mut tx.manifest.metadata, "sources")?,
                    &sid,
                    "Data source",
                )?;
            }
            _ => unreachable!(),
        }
    } else {
        let sheet = tx.sheet_mut(sheet_id)?;
        let regions = ensure_array(&mut sheet.metadata, "dataRegions")?;
        match id {
            "dataRegion.add" => {
                let region = object_clone(
                    p.get("region")
                        .ok_or_else(|| validation("region is required"))?,
                )?;
                same_sheet(&region, sheet_id)?;
                let rid = id_of(&region)?.to_owned();
                ranges.push(range_any(
                    region
                        .get("range")
                        .ok_or_else(|| validation("region range is required"))?,
                )?);
                if find_index(regions, &rid).is_some() {
                    return Err(error(
                        "CONFLICT",
                        format!("Sheet data region already exists: {rid}"),
                    ));
                }
                regions.push(region);
            }
            "dataRegion.remove" => {
                let rid = required_string(p, "regionId")?;
                let i = find_index(regions, &rid).ok_or_else(|| {
                    error("NOT_FOUND", format!("Sheet data region not found: {rid}"))
                })?;
                ranges.push(range_any(
                    regions[i]
                        .get("range")
                        .ok_or_else(|| validation("region range is required"))?,
                )?);
                regions.remove(i);
            }
            _ => unreachable!(),
        }
    }
    for range in ranges {
        affected(tx, range);
    }
    Ok(())
}

fn source_range(tx: &Transaction, source: &Value, ranges: &mut Vec<RangeRef>) -> KernelResult<()> {
    if let Some(raw) = source.get("sourceRange").filter(|v| !v.is_null()) {
        ranges.push(range_any(raw)?);
    } else if let Some(sheet_id) = source.get("sourceSheetId").and_then(Value::as_str) {
        ranges.push(whole_sheet(tx, sheet_id)?);
    }
    Ok(())
}

fn pivot(tx: &mut Transaction, id: &str, sheet_id: &str, p: &Value) -> KernelResult<()> {
    if id == "pivot.drilldown.add" {
        return pivot_drilldown_add(tx, sheet_id, p);
    }
    if id == "pivot.drilldown.remove" {
        return pivot_drilldown_remove(tx, p);
    }
    let mut ranges = Vec::new();
    match id {
        "pivot.add" => {
            let pivot = object_clone(p.get("pivot").unwrap_or(p))?;
            same_sheet(&pivot, sheet_id)?;
            require_pivot_target(&pivot, sheet_id)?;
            let pid = id_of(&pivot)?.to_owned();
            pivot_ranges(tx, &pivot, &mut ranges)?;
            if tx.manifest.sheets.iter().any(|s| {
                s.metadata
                    .get("pivots")
                    .and_then(Value::as_array)
                    .is_some_and(|ps| find_index(ps, &pid).is_some())
            }) {
                return Err(error("CONFLICT", format!("Pivot already exists: {pid}")));
            }
            let sheet = tx.sheet_mut(sheet_id)?;
            let pivots = ensure_array(&mut sheet.metadata, "pivots")?;
            pivots.push(pivot);
        }
        "pivot.remove" => {
            let pid = required_string(p, "pivotId")?;
            let current = tx
                .sheet(sheet_id)?
                .metadata
                .get("pivots")
                .and_then(Value::as_array)
                .and_then(|v| find_index(v, &pid).map(|i| v[i].clone()))
                .ok_or_else(|| error("NOT_FOUND", format!("Pivot not found: {pid}")))?;
            require_pivot_target(&current, sheet_id)?;
            pivot_ranges(tx, &current, &mut ranges)?;
            pivot_dependencies(tx, &pid)?;
            let sheet = tx.sheet_mut(sheet_id)?;
            remove_id(ensure_array(&mut sheet.metadata, "pivots")?, &pid, "Pivot")?;
        }
        "pivot.update" => {
            let pid = required_string(p, "pivotId")?;
            let current = tx
                .sheet(sheet_id)?
                .metadata
                .get("pivots")
                .and_then(Value::as_array)
                .and_then(|v| find_index(v, &pid).map(|i| v[i].clone()))
                .ok_or_else(|| error("NOT_FOUND", format!("Pivot not found: {pid}")))?;
            let mut next = current;
            let next_object = next
                .as_object_mut()
                .ok_or_else(|| validation("Pivot must be an object"))?;
            for k in [
                "source",
                "target",
                "fieldCatalog",
                "layout",
                "refreshPolicy",
                "presentation",
                "nativeMetadata",
            ] {
                if let Some(v) = p.get(k) {
                    next_object.insert(k.into(), v.clone());
                }
            }
            require_pivot_target(&next, sheet_id)?;
            pivot_ranges(tx, &next, &mut ranges)?;
            let sheet = tx.sheet_mut(sheet_id)?;
            let pivots = ensure_array(&mut sheet.metadata, "pivots")?;
            let pos = find_index(pivots, &pid)
                .ok_or_else(|| error("NOT_FOUND", format!("Pivot not found: {pid}")))?;
            pivots[pos] = next;
        }
        "pivot.refresh" => {
            let pid = required_string(p, "pivotId")?;
            let current = tx
                .sheet(sheet_id)?
                .metadata
                .get("pivots")
                .and_then(Value::as_array)
                .and_then(|v| {
                    v.iter()
                        .find(|x| x.get("id").and_then(Value::as_str) == Some(pid.as_str()))
                        .cloned()
                })
                .ok_or_else(|| error("NOT_FOUND", format!("Pivot not found: {pid}")))?;
            require_pivot_target(&current, sheet_id)?;
            pivot_ranges(tx, &current, &mut ranges)?;
        }
        _ => unreachable!(),
    }
    for range in ranges {
        affected(tx, range);
    }
    Ok(())
}

fn require_pivot_target(pivot: &Value, sheet_id: &str) -> KernelResult<()> {
    let target = pivot
        .get("target")
        .and_then(Value::as_object)
        .ok_or_else(|| validation("Pivot target is required"))?;
    if target.get("sheetId").and_then(Value::as_str) != Some(sheet_id) {
        return Err(validation(
            "Pivot target sheetId does not match mutation sheet",
        ));
    }
    Ok(())
}

fn pivot_ranges(tx: &Transaction, pivot: &Value, ranges: &mut Vec<RangeRef>) -> KernelResult<()> {
    let target = pivot
        .get("target")
        .and_then(Value::as_object)
        .ok_or_else(|| validation("Pivot target is required"))?;
    let target_sheet = target.get("sheetId").and_then(Value::as_str).unwrap_or("");
    if target_sheet.is_empty() {
        return Err(validation("Pivot target sheetId is required"));
    }
    ranges.push(whole_sheet(tx, target_sheet)?);
    let source = pivot
        .get("source")
        .and_then(Value::as_object)
        .ok_or_else(|| validation("Pivot source is required"))?;
    match source
        .get("kind")
        .and_then(Value::as_str)
        .unwrap_or_default()
    {
        "worksheet-range" => ranges.push(range_any(
            source
                .get("range")
                .ok_or_else(|| validation("Pivot source range is required"))?,
        )?),
        "worksheet-ranges" => {
            for entry in source
                .get("ranges")
                .and_then(Value::as_array)
                .ok_or_else(|| validation("Pivot source ranges are required"))?
            {
                ranges.push(range_any(
                    entry
                        .get("range")
                        .ok_or_else(|| validation("Pivot source range is required"))?,
                )?);
            }
        }
        "table" | "named-range" | "data-source" => {}
        _ => return Err(validation("Pivot source kind is invalid")),
    }
    Ok(())
}
fn pivot_dependencies(tx: &Transaction, pivot_id: &str) -> KernelResult<()> {
    for sheet in &tx.manifest.sheets {
        let drawings = sheet
            .metadata
            .get("drawings")
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default();
        let payloads = sheet
            .metadata
            .get("drawingPayloads")
            .and_then(Value::as_object)
            .cloned()
            .unwrap_or_default();
        for drawing in drawings {
            let payload_id = drawing
                .get("payloadId")
                .and_then(Value::as_str)
                .unwrap_or_default();
            if let Some(payload) = payloads.get(payload_id) {
                let source = if payload.get("kind").and_then(Value::as_str) == Some("chart") {
                    payload.get("source").unwrap_or(payload)
                } else {
                    payload
                };
                let direct = matches!(
                    payload.get("kind").and_then(Value::as_str),
                    Some("chart" | "slicer" | "timeline")
                ) && source.get("pivotId").and_then(Value::as_str) == Some(pivot_id);
                let connected = matches!(
                    payload.get("kind").and_then(Value::as_str),
                    Some("slicer" | "timeline")
                ) && payload
                    .get("connections")
                    .and_then(Value::as_array)
                    .is_some_and(|xs| {
                        xs.iter()
                            .any(|x| x.get("pivotId").and_then(Value::as_str) == Some(pivot_id))
                    });
                if direct || connected {
                    return Err(error("CONFLICT", "Pivot has dependent drawing"));
                }
            }
        }
    }
    Ok(())
}

fn pivot_drilldown_add(tx: &mut Transaction, pivot_sheet_id: &str, p: &Value) -> KernelResult<()> {
    let pivot_id = required_string(p, "pivotId")?;
    let label = required_string(p, "label")?;
    let target_sheet_id = required_string(p, "targetSheetId")?;
    let target = required_object(p, "target")?;
    let anchor_row = target
        .get("row")
        .and_then(Value::as_u64)
        .and_then(|v| u32::try_from(v).ok())
        .ok_or_else(|| validation("Pivot drill-down target row is invalid"))?;
    let anchor_column = target
        .get("column")
        .and_then(Value::as_u64)
        .and_then(|v| u32::try_from(v).ok())
        .ok_or_else(|| validation("Pivot drill-down target column is invalid"))?;
    let pivot = tx
        .sheet(pivot_sheet_id)?
        .metadata
        .get("pivots")
        .and_then(Value::as_array)
        .and_then(|v| find_index(v, &pivot_id).map(|i| v[i].clone()))
        .ok_or_else(|| error("NOT_FOUND", format!("Pivot not found: {pivot_id}")))?;
    if pivot.get("schema").and_then(Value::as_str) != Some("PivotDefinition") {
        return Err(validation("Pivot schema must be PivotDefinition"));
    }
    if tx
        .manifest
        .sheets
        .iter()
        .any(|s| s.sheet_id == target_sheet_id)
    {
        return Err(error(
            "CONFLICT",
            format!("Pivot drill-down target already exists: {target_sheet_id}"),
        ));
    }
    let source_ranges = pivot_worksheet_ranges(&pivot)?;
    if source_ranges.is_empty() {
        return Err(validation("Pivot drill-down source has no worksheet range"));
    }
    let paths = required_array(p, "sourceRowPaths")?;
    if paths.len() > 1_048_576 {
        return Err(validation("Pivot drill-down has too many source rows"));
    }
    let mut source_paths = Vec::with_capacity(paths.len());
    for path in paths {
        let source_sheet = required_string(path, "sheetId")?;
        let row = required_u32(path, "row")?;
        let sheet = tx.sheet(&source_sheet)?;
        if row >= sheet.row_count {
            return Err(validation("Pivot drill-down source row exceeds worksheet"));
        }
        if !source_ranges
            .iter()
            .any(|r| r.sheet_id == source_sheet && row > r.start_row && row <= r.end_row)
        {
            return Err(validation(
                "Pivot drill-down row is outside its canonical source",
            ));
        }
        source_paths.push((source_sheet, row));
    }
    let mut columns = Vec::new();
    for range in &source_ranges {
        for column in range.start_column..=range.end_column {
            let label = tx
                .read(&range.sheet_id, range.start_row, column)?
                .map(|c| scalar_label(&c.value))
                .filter(|s| !s.is_empty())
                .unwrap_or_else(|| format!("Column {}", column - range.start_column + 1));
            columns.push((range.clone(), column, label));
        }
    }
    if columns.is_empty() {
        return Err(validation("Pivot drill-down source has no columns"));
    }
    let rows_per_result = source_ranges.len().max(1);
    if source_paths.len() % rows_per_result != 0 {
        return Err(validation(
            "Pivot drill-down source paths do not form complete source rows",
        ));
    }
    let detail_rows = (source_paths.len() + rows_per_result - 1) / rows_per_result;
    let row_count = anchor_row
        .checked_add(
            u32::try_from(detail_rows).map_err(|_| validation("Drill-down row count overflows"))?,
        )
        .and_then(|n| n.checked_add(1))
        .filter(|n| *n <= kernel_core::MAX_ROWS)
        .ok_or_else(|| validation("Pivot drill-down target exceeds worksheet rows"))?
        .max(1_000);
    let column_count = anchor_column
        .checked_add(
            u32::try_from(columns.len())
                .map_err(|_| validation("Drill-down column count overflows"))?,
        )
        .filter(|n| *n <= kernel_core::MAX_COLUMNS)
        .ok_or_else(|| validation("Pivot drill-down target exceeds worksheet columns"))?
        .max(26);
    let name = format!("Drill {pivot_id} {label}");
    let name = name.chars().take(31).collect::<String>();
    let metadata = drilldown_sheet_metadata();
    tx.manifest.sheets.push(kernel_core::SheetManifest {
        sheet_id: target_sheet_id.clone(),
        name,
        row_count,
        column_count,
        metadata,
    });
    for (column_offset, (_, _, label)) in columns.iter().enumerate() {
        tx.write(
            &target_sheet_id,
            anchor_row,
            anchor_column + column_offset as u32,
            Some(Cell {
                value: Scalar::Text(label.clone()),
                formula: None,
                metadata: BTreeMap::new(),
            }),
        )?;
    }
    for row_offset in 0..detail_rows {
        let start = row_offset * rows_per_result;
        let end = (start + rows_per_result).min(source_paths.len());
        let selected = &source_paths[start..end];
        for (column_offset, (range, column, _)) in columns.iter().enumerate() {
            let value = selected
                .iter()
                .find(|(sheet, row)| {
                    sheet == &range.sheet_id && *row > range.start_row && *row <= range.end_row
                })
                .map(|(_, row)| tx.read(&range.sheet_id, *row, *column))
                .transpose()?
                .flatten()
                .map(|cell| cell.value)
                .unwrap_or(Scalar::Null);
            tx.write(
                &target_sheet_id,
                anchor_row + row_offset as u32 + 1,
                anchor_column + column_offset as u32,
                Some(Cell {
                    value,
                    formula: None,
                    metadata: BTreeMap::new(),
                }),
            )?;
        }
    }
    for range in source_ranges {
        affected(tx, range);
    }
    affected(
        tx,
        RangeRef {
            sheet_id: target_sheet_id,
            start_row: 0,
            end_row: row_count - 1,
            start_column: 0,
            end_column: column_count - 1,
        },
    );
    Ok(())
}

fn pivot_drilldown_remove(tx: &mut Transaction, p: &Value) -> KernelResult<()> {
    let target_sheet_id = required_string(p, "targetSheetId")?;
    if tx.manifest.sheets.len() <= 1 {
        return Err(validation("A workbook must keep at least one worksheet"));
    }
    let index = tx
        .manifest
        .sheets
        .iter()
        .position(|s| s.sheet_id == target_sheet_id)
        .ok_or_else(|| {
            error(
                "NOT_FOUND",
                format!("Pivot drill-down target not found: {target_sheet_id}"),
            )
        })?;
    let current = tx.manifest.sheets[index].clone();
    if !current.name.starts_with("Drill ") {
        return Err(error(
            "FORBIDDEN",
            "Only a server-created pivot drill-down sheet may be removed",
        ));
    }
    affected(
        tx,
        RangeRef {
            sheet_id: target_sheet_id.clone(),
            start_row: 0,
            end_row: current.row_count - 1,
            start_column: 0,
            end_column: current.column_count - 1,
        },
    );
    tx.manifest.sheets.remove(index);
    if let Some(documents) = tx
        .manifest
        .metadata
        .get_mut("printDocuments")
        .and_then(Value::as_array_mut)
    {
        documents
            .retain(|v| v.get("sheetId").and_then(Value::as_str) != Some(target_sheet_id.as_str()));
    }
    if let Some(names) = tx
        .manifest
        .metadata
        .get_mut("definedNameModels")
        .and_then(Value::as_array_mut)
    {
        names.retain(|v| {
            !(v.get("scope").and_then(Value::as_str) == Some("sheet")
                && v.get("sheetId").and_then(Value::as_str) == Some(target_sheet_id.as_str()))
        });
    }
    Ok(())
}

fn pivot_worksheet_ranges(pivot: &Value) -> KernelResult<Vec<RangeRef>> {
    let source = pivot
        .get("source")
        .and_then(Value::as_object)
        .ok_or_else(|| validation("Pivot source is required"))?;
    match source
        .get("kind")
        .and_then(Value::as_str)
        .unwrap_or_default()
    {
        "worksheet-range" => {
            Ok(vec![range_any(source.get("range").ok_or_else(|| {
                validation("Pivot source range is required")
            })?)?])
        }
        "worksheet-ranges" => source
            .get("ranges")
            .and_then(Value::as_array)
            .ok_or_else(|| validation("Pivot source ranges are required"))?
            .iter()
            .map(|entry| {
                range_any(
                    entry
                        .get("range")
                        .ok_or_else(|| validation("Pivot source range is required"))?,
                )
            })
            .collect(),
        _ => Err(validation("Pivot drill-down source has no worksheet range")),
    }
}

fn scalar_label(value: &Scalar) -> String {
    match value {
        Scalar::Null => String::new(),
        Scalar::Boolean(v) => v.to_string(),
        Scalar::Number(v) => v.to_string(),
        Scalar::Text(v) => v.clone(),
        Scalar::Error(v) => v.code.clone(),
    }
}

fn drilldown_sheet_metadata() -> BTreeMap<String, Value> {
    let mut metadata = BTreeMap::new();
    for key in [
        "pivots",
        "sparklines",
        "sparklineGroups",
        "drawings",
        "conditionalFormats",
        "dataValidations",
        "hiddenRows",
        "hiddenColumns",
        "sheetTables",
        "protectionRules",
        "dataRegions",
        "hyperlinks",
    ] {
        metadata.insert(key.to_owned(), Value::Array(Vec::new()));
    }
    metadata.insert("drawingPayloads".into(), Value::Object(Map::new()));
    metadata.insert("showGridlines".into(), Value::Bool(true));
    metadata.insert("showHeaders".into(), Value::Bool(true));
    metadata
}

fn query(tx: &mut Transaction, id: &str, sheet_id: &str, p: &Value) -> KernelResult<()> {
    if id == "query.definition.replace" {
        let qid = required_string(p, "queryId")?;
        let defs = ensure_array(&mut tx.manifest.metadata, "queryDefinitions")?;
        match p
            .get("definition")
            .ok_or_else(|| validation("Query definition is required"))?
        {
            Value::Null => {
                if let Some(i) = find_index(defs, &qid) {
                    defs.remove(i);
                }
            }
            v => {
                let d = object_clone(v)?;
                if d.get("schema").and_then(Value::as_str) != Some("QueryDefinition")
                    || d.get("id").and_then(Value::as_str) != Some(qid.as_str())
                {
                    return Err(validation("Query definition schema or id is invalid"));
                }
                if let Some(i) = find_index(defs, &qid) {
                    defs[i] = d
                } else {
                    defs.push(d)
                }
            }
        }
        return Ok(());
    }
    let qid = required_string(p, "queryId")?;
    let source_id = required_string(p, "sourceId")?;
    if source_id != format!("query:{qid}") {
        return Err(validation("Query sourceId does not match queryId"));
    }
    let target = required_object(p, "target")?;
    let target_kind = target
        .get("kind")
        .and_then(Value::as_str)
        .ok_or_else(|| validation("Query load target kind is required"))?;
    if !matches!(
        target_kind,
        "range" | "sheet-table" | "pivot-source" | "workbook-table"
    ) {
        return Err(validation("Query load target kind is invalid"));
    }
    match target_kind {
        "range" | "sheet-table" => {
            let target_sheet = target
                .get("sheetId")
                .and_then(Value::as_str)
                .ok_or_else(|| validation("Query load target sheetId is required"))?;
            tx.sheet(target_sheet)?;
        }
        "workbook-table" => {
            let table_id = target
                .get("tableId")
                .and_then(Value::as_str)
                .ok_or_else(|| validation("Query load target tableId is required"))?;
            if !tx
                .manifest
                .metadata
                .get("tables")
                .and_then(Value::as_array)
                .is_some_and(|ts| find_index(ts, table_id).is_some())
            {
                return Err(error(
                    "NOT_FOUND",
                    format!("Workbook table not found: {table_id}"),
                ));
            }
        }
        "pivot-source" => {
            let pivot_id = target
                .get("pivotId")
                .and_then(Value::as_str)
                .ok_or_else(|| validation("Query load target pivotId is required"))?;
            if !tx.manifest.sheets.iter().any(|s| {
                s.metadata
                    .get("pivots")
                    .and_then(Value::as_array)
                    .is_some_and(|ps| find_index(ps, pivot_id).is_some())
            }) {
                return Err(error("NOT_FOUND", format!("Pivot not found: {pivot_id}")));
            }
        }
        _ => unreachable!(),
    }
    // Query execution is server-owned. A load without its sealed result is
    // rejected so a client cannot manufacture rows or silently persist only a
    // header.
    let (result_columns, result_rows) = trusted_result(p)?;
    // A load replaces the existing binding for this source.  Capture ranges
    // before mutating metadata so stale query rows cannot survive a reload.
    let mut old_ranges = Vec::new();
    for s in &tx.manifest.sheets {
        if let Some(regions) = s.metadata.get("dataRegions").and_then(Value::as_array) {
            for region in regions {
                if region.get("sourceId").and_then(Value::as_str) == Some(source_id.as_str()) {
                    old_ranges.push(range_any(
                        region
                            .get("range")
                            .ok_or_else(|| validation("Query region range is required"))?,
                    )?);
                }
            }
        }
    }
    for r in &old_ranges {
        affected(tx, r.clone());
        for row in r.start_row..=r.end_row {
            for col in r.start_column..=r.end_column {
                if tx.read(&r.sheet_id, row, col)?.is_some() {
                    tx.write(&r.sheet_id, row, col, None)?;
                }
            }
        }
    }
    for sheet in &mut tx.manifest.sheets {
        if let Some(regions) = sheet
            .metadata
            .get_mut("dataRegions")
            .and_then(Value::as_array_mut)
        {
            regions
                .retain(|r| r.get("sourceId").and_then(Value::as_str) != Some(source_id.as_str()));
        }
    }
    if let Some(tables) = tx
        .manifest
        .metadata
        .get_mut("tables")
        .and_then(Value::as_array_mut)
    {
        for table in tables {
            if table.get("sourceId").and_then(Value::as_str) == Some(source_id.as_str()) {
                if let Some(o) = table.as_object_mut() {
                    o.remove("sourceId");
                }
            }
        }
    }
    if let Some(sources) = tx
        .manifest
        .metadata
        .get_mut("sources")
        .and_then(Value::as_array_mut)
    {
        if let Some(i) = find_index(sources, &source_id) {
            sources.remove(i);
        }
    }
    let defs = ensure_array(&mut tx.manifest.metadata, "queryDefinitions")?;
    if let Some(d) = p.get("queryDefinition").filter(|v| !v.is_null()) {
        let d = object_clone(d)?;
        if let Some(i) = find_index(defs, &qid) {
            defs[i] = d
        } else {
            defs.push(d)
        }
    } else if let Some(i) = find_index(defs, &qid) {
        defs.remove(i);
    }
    let sources = ensure_array(&mut tx.manifest.metadata, "sources")?;
    if let Some(source) = p.get("source").filter(|v| !v.is_null()) {
        let source = object_clone(source)?;
        if let Some(source_sheet) = source.get("sourceSheetId").and_then(Value::as_str) {
            if source_sheet != sheet_id {
                return Err(validation(
                    "Query data source sourceSheetId does not match mutation sheet",
                ));
            }
        }
        if let Some(i) = find_index(sources, &source_id) {
            sources[i] = source
        } else {
            sources.push(source)
        }
    }
    if let Some(extent) = p.get("extent").filter(|v| !v.is_null()) {
        let o = extent
            .as_object()
            .ok_or_else(|| validation("Query load extent must be an object"))?;
        let target_sheet = o
            .get("sheetId")
            .and_then(Value::as_str)
            .ok_or_else(|| validation("Query load extent sheetId is required"))?
            .to_owned();
        let rows = o
            .get("rowCount")
            .and_then(Value::as_u64)
            .ok_or_else(|| validation("Query load extent rowCount is invalid"))?;
        let cols = o
            .get("columnCount")
            .and_then(Value::as_u64)
            .ok_or_else(|| validation("Query load extent columnCount is invalid"))?;
        if rows == 0 || cols == 0 {
            return Err(validation("Query load extent must be positive"));
        }
        let target = tx.sheet_mut(&target_sheet)?;
        target.row_count = u32::try_from(rows)
            .map_err(|_| validation("Query load extent rowCount is too large"))?;
        target.column_count = u32::try_from(cols)
            .map_err(|_| validation("Query load extent columnCount is too large"))?;
    }
    if let Some(binding) = p.get("binding").filter(|v| !v.is_null()) {
        let kind = binding
            .get("kind")
            .and_then(Value::as_str)
            .unwrap_or_default();
        if kind == "sheet-region" {
            let region = binding
                .get("region")
                .ok_or_else(|| validation("Query sheet region is required"))?
                .clone();
            let r = range_any(
                region
                    .get("range")
                    .ok_or_else(|| validation("Query region range is required"))?,
            )?;
            tx.sheet(&r.sheet_id)?;
            affected(tx, r.clone());
            {
                let sheet = tx.sheet_mut(&r.sheet_id)?;
                let regions = ensure_array(&mut sheet.metadata, "dataRegions")?;
                let rid = required_string(&region, "id")?;
                if let Some(i) = find_index(regions, &rid) {
                    regions[i] = region
                } else {
                    regions.push(region)
                }
            }
        } else if kind == "workbook-table" {
            if let Some(table) = binding.get("table") {
                let tables = ensure_array(&mut tx.manifest.metadata, "tables")?;
                let tid = table.get("id").and_then(Value::as_str).unwrap_or_default();
                if let Some(i) = find_index(tables, tid) {
                    tables[i] = table.clone()
                }
            }
        } else {
            return Err(validation("Query load binding kind is invalid"));
        }
    }
    if let Some(pivot_source) = p.get("pivotSource").filter(|v| !v.is_null()) {
        let target = p
            .get("target")
            .and_then(Value::as_object)
            .ok_or_else(|| validation("Query load target is required"))?;
        let pivot_id = target
            .get("pivotId")
            .and_then(Value::as_str)
            .filter(|v| !v.is_empty())
            .ok_or_else(|| validation("Query load pivotId is required"))?;
        let source = object_clone(pivot_source)?;
        match source
            .get("kind")
            .and_then(Value::as_str)
            .unwrap_or_default()
        {
            "data-source" => {
                if source.get("dataSourceId").and_then(Value::as_str) != Some(source_id.as_str()) {
                    return Err(validation(
                        "Query Pivot sourceId does not match query source",
                    ));
                }
            }
            "worksheet-range" => {
                let _ = range_any(
                    source
                        .get("range")
                        .ok_or_else(|| validation("Query Pivot source range is required"))?,
                )?;
            }
            "worksheet-ranges" => {
                if source.get("ranges").and_then(Value::as_array).is_none() {
                    return Err(validation("Query Pivot worksheet sources are invalid"));
                }
            }
            "table" => {
                required_string(&source, "tableId")?;
            }
            "named-range" => {
                required_string(&source, "name")?;
            }
            _ => return Err(validation("Query Pivot source kind is invalid")),
        }
        let mut found = false;
        for sheet in &mut tx.manifest.sheets {
            if let Some(pivots) = sheet
                .metadata
                .get_mut("pivots")
                .and_then(Value::as_array_mut)
            {
                if let Some(i) = find_index(pivots, pivot_id) {
                    pivots[i]
                        .as_object_mut()
                        .ok_or_else(|| validation("Pivot must be an object"))?
                        .insert("source".into(), source.clone());
                    found = true;
                }
            }
        }
        if !found {
            return Err(error("NOT_FOUND", format!("Pivot not found: {pivot_id}")));
        }
    }
    if let Some(destination) = query_destination(
        tx,
        p,
        target_kind,
        result_columns.len(),
        result_rows.len() + 1,
    )? {
        write_query_result(tx, &destination, &result_columns, &result_rows)?;
    } else if target_kind == "range" {
        return Err(validation("Query range load target range is required"));
    }
    Ok(())
}

fn trusted_result(p: &Value) -> KernelResult<(Vec<String>, Vec<Vec<Value>>)> {
    let result = p
        .get("trustedResult")
        .and_then(Value::as_object)
        .ok_or_else(|| validation("Query load trustedResult is required"))?;
    let columns = result
        .get("columns")
        .and_then(Value::as_array)
        .ok_or_else(|| validation("Query load trustedResult.columns is required"))?;
    if columns.is_empty() || columns.len() > 16_384 {
        return Err(validation("Query load trustedResult columns are invalid"));
    }
    let columns = columns
        .iter()
        .map(|v| {
            v.as_str()
                .filter(|s| !s.is_empty())
                .map(str::to_owned)
                .ok_or_else(|| validation("Query load trustedResult column is invalid"))
        })
        .collect::<KernelResult<Vec<_>>>()?;
    let rows = result
        .get("rows")
        .and_then(Value::as_array)
        .ok_or_else(|| validation("Query load trustedResult.rows is required"))?;
    let mut out = Vec::with_capacity(rows.len());
    for row in rows {
        let values = row
            .as_array()
            .ok_or_else(|| validation("Query load trustedResult row is invalid"))?;
        if values.len() != columns.len() {
            return Err(validation(
                "Query load trustedResult row width does not match columns",
            ));
        }
        for value in values {
            scalar_value(value)?;
        }
        out.push(values.clone());
    }
    Ok((columns, out))
}

fn query_destination(
    tx: &Transaction,
    p: &Value,
    target_kind: &str,
    width: usize,
    height: usize,
) -> KernelResult<Option<RangeRef>> {
    if let Some(binding) = p.get("binding").filter(|v| !v.is_null()) {
        if binding.get("kind").and_then(Value::as_str) == Some("sheet-region") {
            let region = binding
                .get("region")
                .and_then(Value::as_object)
                .ok_or_else(|| validation("Query sheet region is required"))?;
            return Ok(Some(range_any(
                region
                    .get("range")
                    .ok_or_else(|| validation("Query region range is required"))?,
            )?));
        }
    }
    if target_kind != "range" {
        return Ok(None);
    }
    let target = required_object(p, "target")?;
    let sheet_id = target
        .get("sheetId")
        .and_then(Value::as_str)
        .ok_or_else(|| validation("Query range target sheetId is required"))?;
    let raw = target
        .get("range")
        .and_then(Value::as_object)
        .ok_or_else(|| validation("Query range target range is required"))?;
    let start_row = number_u32(raw, "startRow")?;
    let start_column = number_u32(raw, "startColumn")?;
    let end_row = raw
        .get("endRow")
        .map(|_| number_u32(raw, "endRow"))
        .transpose()?
        .unwrap_or_else(|| start_row.saturating_add(height.saturating_sub(1) as u32));
    let end_column = raw
        .get("endColumn")
        .map(|_| number_u32(raw, "endColumn"))
        .transpose()?
        .unwrap_or_else(|| start_column.saturating_add(width.saturating_sub(1) as u32));
    let range = RangeRef {
        sheet_id: sheet_id.to_owned(),
        start_row,
        end_row,
        start_column,
        end_column,
    };
    range.validate()?;
    tx.sheet(sheet_id)?;
    Ok(Some(range))
}

fn write_query_result(
    tx: &mut Transaction,
    destination: &RangeRef,
    columns: &[String],
    rows: &[Vec<Value>],
) -> KernelResult<()> {
    if columns.len() as u32 > destination.end_column - destination.start_column + 1
        || rows.len() as u32 + 1 > destination.end_row - destination.start_row + 1
    {
        return Err(validation("Query result exceeds target range"));
    }
    let existing = tx.cells(&destination.sheet_id)?;
    for (address, _) in existing {
        if destination.contains(&address) {
            tx.write(&address.sheet_id, address.row, address.column, None)?;
        }
    }
    for (offset, column) in columns.iter().enumerate() {
        tx.write(
            &destination.sheet_id,
            destination.start_row,
            destination.start_column + offset as u32,
            Some(Cell {
                value: Scalar::Text(column.clone()),
                formula: None,
                metadata: Default::default(),
            }),
        )?;
    }
    for (row_offset, row) in rows.iter().enumerate() {
        for (column_offset, value) in row.iter().enumerate() {
            tx.write(
                &destination.sheet_id,
                destination.start_row + row_offset as u32 + 1,
                destination.start_column + column_offset as u32,
                Some(Cell {
                    value: scalar_value(value)?,
                    formula: None,
                    metadata: Default::default(),
                }),
            )?;
        }
    }
    Ok(())
}

fn scalar_value(v: &Value) -> KernelResult<Scalar> {
    if !(v.is_null() || v.is_boolean() || v.is_number() || v.is_string()) {
        return Err(validation("Query header contains a non-scalar value"));
    }
    serde_json::from_value(v.clone())
        .map_err(|_| validation("Query header contains an invalid scalar"))
}

fn workbook_state(tx: &mut Transaction, id: &str, sheet_id: &str, p: &Value) -> KernelResult<()> {
    match id {
        "table.add" => {
            let table = object_clone(p.get("table").unwrap_or(p))?;
            if let Some(raw) = table.get("sourceRange").filter(|v| !v.is_null()) {
                affected(tx, range_any(raw)?);
            }
            let tables = ensure_array(&mut tx.manifest.metadata, "tables")?;
            let tid = id_of(&table)?.to_owned();
            if find_index(tables, &tid).is_some() {
                return Err(error(
                    "CONFLICT",
                    format!("Workbook table already exists: {tid}"),
                ));
            }
            tables.push(table);
        }
        "table.remove" => {
            let tid = required_string(p, "tableId")?;
            let existing = tx
                .manifest
                .metadata
                .get("tables")
                .and_then(Value::as_array)
                .and_then(|v| find_index(v, &tid).map(|i| v[i].clone()))
                .ok_or_else(|| error("NOT_FOUND", format!("Workbook table not found: {tid}")))?;
            if let Some(raw) = existing.get("sourceRange").filter(|v| !v.is_null()) {
                affected(tx, range_any(raw)?);
            }
            remove_id(
                ensure_array(&mut tx.manifest.metadata, "tables")?,
                &tid,
                "Workbook table",
            )?;
        }
        "name.set" => {
            let model = object_clone(
                p.get("model")
                    .ok_or_else(|| validation("model is required"))?,
            )?;
            let name = model
                .get("name")
                .and_then(Value::as_str)
                .ok_or_else(|| validation("defined name is invalid"))?;
            let scope = model
                .get("scope")
                .and_then(Value::as_str)
                .unwrap_or("workbook");
            let sheet = model.get("sheetId").and_then(Value::as_str);
            let names = ensure_array(&mut tx.manifest.metadata, "definedNameModels")?;
            let i = names.iter().position(|v| {
                v.get("name")
                    .and_then(Value::as_str)
                    .is_some_and(|n| n.eq_ignore_ascii_case(name))
                    && v.get("scope").and_then(Value::as_str) == Some(scope)
                    && v.get("sheetId").and_then(Value::as_str) == sheet
            });
            if let Some(i) = i {
                names[i] = model
            } else {
                names.push(model)
            }
        }
        "name.remove" => {
            let name = required_string(p, "name")?;
            let scope = p.get("scope").and_then(Value::as_str).unwrap_or("workbook");
            let sheet = p.get("sheetId").and_then(Value::as_str);
            let names = ensure_array(&mut tx.manifest.metadata, "definedNameModels")?;
            let i = names
                .iter()
                .position(|v| {
                    v.get("name")
                        .and_then(Value::as_str)
                        .is_some_and(|n| n.eq_ignore_ascii_case(&name))
                        && v.get("scope").and_then(Value::as_str) == Some(scope)
                        && v.get("sheetId").and_then(Value::as_str) == sheet
                })
                .ok_or_else(|| error("NOT_FOUND", format!("Defined name not found: {name}")))?;
            names.remove(i);
        }
        "workbook.calculation.mode.set" => {
            let mode = required_string(p, "mode")?;
            if !matches!(mode.as_str(), "automatic" | "manual" | "partial") {
                return Err(validation("Workbook calculation mode is invalid"));
            }
            let settings = ensure_object(&mut tx.manifest.metadata, "calculationSettings")?;
            settings.insert("mode".into(), json!(mode));
        }
        _ => unreachable!(),
    }
    Ok(())
}
