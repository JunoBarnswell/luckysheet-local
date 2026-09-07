use crate::{Transaction, invalid, range, text};
use kernel_core::*;
use serde_json::{Value, json};

pub(crate) fn apply(tx: &mut Transaction, id: &str, sheet: &str, p: &Value) -> KernelResult<bool> {
    if !matches!(id, "sheet.protect.set" | "sheet.protect.remove") {
        return Ok(false);
    }
    let mut rules = tx
        .sheet(sheet)?
        .metadata
        .get("protectionRules")
        .cloned()
        .unwrap_or_else(|| json!([]));
    let rules = rules
        .as_array_mut()
        .ok_or_else(|| invalid("Protection rules must be array"))?;
    if id == "sheet.protect.set" {
        let rule = p.get("rule").ok_or_else(|| invalid("rule is required"))?;
        let rid = text(rule, "id")?;
        let scope = text(rule, "scope")?;
        if !["workbook", "sheet", "range"].contains(&scope) || !rule["locked"].is_boolean() {
            return Err(invalid("Invalid protection rule"));
        }
        if scope != "workbook" && text(rule, "sheetId")? != sheet {
            return Err(invalid("Protection worksheet mismatch"));
        }
        if scope == "range" {
            range(tx, sheet, &rule["range"])?;
        }
        if let Some(allow) = rule.get("allow") {
            let allow = allow
                .as_object()
                .ok_or_else(|| invalid("Protection allow must be object"))?;
            for (key, v) in allow {
                if ![
                    "formatCells",
                    "insertRows",
                    "insertColumns",
                    "deleteRows",
                    "deleteColumns",
                    "sort",
                    "autoFilter",
                    "editObjects",
                    "selectLocked",
                    "selectUnlocked",
                ]
                .contains(&key.as_str())
                    || !v.is_boolean()
                {
                    return Err(invalid("Invalid protection allow field"));
                }
            }
        }
        if let Some(at) = rules.iter().position(|r| r["id"].as_str() == Some(rid)) {
            rules[at] = rule.clone();
        } else {
            rules.push(rule.clone());
        }
    } else {
        let rid = text(p, "ruleId")?;
        let at = rules
            .iter()
            .position(|r| r["id"].as_str() == Some(rid))
            .ok_or_else(|| KernelError::new("NOT_FOUND", "Protection rule not found"))?;
        rules.remove(at);
    }
    tx.sheet_mut(sheet)?
        .metadata
        .insert("protectionRules".into(), json!(rules));
    Ok(true)
}

pub(crate) fn check(
    tx: &Transaction,
    base: &WorkbookManifest,
    id: &str,
    ranges: &[RangeRef],
) -> KernelResult<()> {
    let action = action(id);
    if action == "none" {
        return Ok(());
    }
    for target in ranges {
        for sheet in &base.sheets {
            let Some(raw) = sheet.metadata.get("protectionRules") else {
                continue;
            };
            for rule in raw
                .as_array()
                .ok_or_else(|| invalid("Protection rules must be array"))?
            {
                if !rule["locked"].is_boolean() {
                    return Err(invalid("Protection locked must be Boolean"));
                }
                if rule["locked"] == false {
                    continue;
                }
                let scope = text(rule, "scope")?;
                let applies = match scope {
                    "workbook" => true,
                    "sheet" => sheet.sheet_id == target.sheet_id,
                    "range" => {
                        let r: RangeRef = serde_json::from_value(rule["range"].clone())
                            .map_err(|_| invalid("Invalid protected range"))?;
                        r.validate()?;
                        r.intersects(target)
                    }
                    _ => return Err(invalid("Unknown protection scope")),
                };
                if !applies {
                    continue;
                }
                if action == "edit-cell" {
                    if scope == "range" {
                        return Err(KernelError::new("FORBIDDEN", "Protected range blocks edit")
                            .at(&target.sheet_id));
                    }
                    for row in target.start_row..=target.end_row {
                        for col in target.start_column..=target.end_column {
                            let address = CellAddress {
                                sheet_id: target.sheet_id.clone(),
                                row,
                                column: col,
                            };
                            let unlocked = tx.base.read_cell(&address)?.is_some_and(|c| {
                                c.metadata
                                    .get("style")
                                    .and_then(|s| s.get("locked"))
                                    .and_then(Value::as_bool)
                                    == Some(false)
                            });
                            if !unlocked {
                                return Err(KernelError::new(
                                    "FORBIDDEN",
                                    "Protected cell blocks edit",
                                )
                                .at(format!("{}:{row}:{col}", target.sheet_id)));
                            }
                        }
                    }
                } else if rule["allow"][action].as_bool() != Some(true) {
                    return Err(KernelError::new(
                        "FORBIDDEN",
                        format!("Protected area blocks {action}"),
                    )
                    .at(&target.sheet_id));
                }
            }
        }
    }
    Ok(())
}
fn action(id: &str) -> &'static str {
    if id.starts_with("comment.")
        || id.starts_with("note.")
        || id.starts_with("sheet.protect.")
        || id.starts_with("workbook.")
        || matches!(
            id,
            "sheet.add"
                | "sheet.remove"
                | "sheet.rename"
                | "sheet.duplicated"
                | "sheet.restore"
                | "sheet.reordered"
                | "sheet.extent.grow"
                | "sheet.extent.restore"
                | "sheet.hidden"
                | "sheet.unhidden"
        )
    {
        return "none";
    }
    if id.starts_with("drawing.") {
        return "editObjects";
    }
    if id.contains("autoFilter") {
        return "autoFilter";
    }
    match id {
        "rows.inserted" => "insertRows",
        "rows.deleted" => "deleteRows",
        "columns.inserted" => "insertColumns",
        "columns.deleted" => "deleteColumns",
        "rows.permuted" => "sort",
        _ => {
            if id.starts_with("style.")
                || id.starts_with("cf.")
                || id.starts_with("dv.")
                || id.starts_with("merge.")
                || id.starts_with("pageLayout.")
                || id.ends_with(".resize")
                || id.starts_with("sheetTable.")
                || matches!(
                    id,
                    "banded.set" | "cell.editor.set" | "sheet.tabColor" | "freeze.set" | "view.set"
                )
            {
                "formatCells"
            } else {
                "edit-cell"
            }
        }
    }
}
