//! Authored entry rules are evaluated against the staged canonical reader.
use crate::{Transaction, invalid, range, text};
use kernel_core::*;
use kernel_formula::{DefinedNameScope, FormulaRuntime, FormulaTable, FormulaValue};
use serde_json::Value;

pub(crate) fn check(
    tx: &Transaction,
    id: &str,
    p: &Value,
    touched: &[CellAddress],
) -> KernelResult<()> {
    if matches!(id, "style.set" | "style.preset.set") {
        if let Some(style) = p.get("style") {
            let s = style
                .as_object()
                .ok_or_else(|| invalid("Style must be object"))?;
            if s.contains_key("unsupportedAlignment") {
                return Err(KernelError::new(
                    "UNSUPPORTED_FEATURE",
                    "Cannot edit preserved alignment attributes",
                ));
            }
            for (key, values) in [
                (
                    "horizontalAlignment",
                    &[
                        "general",
                        "left",
                        "center",
                        "right",
                        "centerContinuous",
                        "justify",
                        "distributed",
                        "fill",
                    ][..],
                ),
                (
                    "verticalAlignment",
                    &["top", "middle", "bottom", "justify", "distributed"][..],
                ),
                ("readingOrder", &["context", "ltr", "rtl"][..]),
                (
                    "textOrientation",
                    &["horizontal", "stacked", "rotateUp", "rotateDown"][..],
                ),
            ] {
                if let Some(v) = s.get(key) {
                    if !v.as_str().is_some_and(|v| values.contains(&v)) {
                        return Err(invalid(format!("Invalid style {key}")));
                    }
                }
            }
            if s.get("shrinkToFit").is_some_and(|v| !v.is_boolean())
                || s.get("indent")
                    .is_some_and(|v| !v.as_u64().is_some_and(|n| n <= 250))
                || s.get("textRotate").is_some_and(|v| {
                    !v.as_f64()
                        .is_some_and(|n| n.is_finite() && (-180.0..=180.0).contains(&n))
                })
            {
                return Err(invalid("Invalid style sizing or orientation"));
            }
        }
        return Ok(());
    }
    if !matches!(
        id,
        "cell.set" | "range.set" | "range.paste" | "fill.applied"
    ) {
        return Ok(());
    }
    if id == "cell.set" {
        if let Some(authority) = p.get("writeAuthority") {
            if ![
                "direct-entry",
                "paste",
                "fill",
                "formula-result",
                "query-load",
                "script",
                "external-sync",
            ]
            .contains(&text(authority, "kind")?)
            {
                return Err(invalid("Unknown cell write intent kind"));
            }
            let target = &authority["target"];
            if target.get("sheetId") != p.get("sheetId")
                || target.get("row") != p.get("row")
                || target.get("column") != p.get("column")
                || authority.get("candidate") != p.get("value")
            {
                return Err(invalid(
                    "Cell write intent does not match its authored candidate",
                ));
            }
            if !["accepted", "confirmed"]
                .contains(&text(&authority["validationDecision"], "status")?)
            {
                return Err(invalid("Invalid cell validation decision"));
            }
        }
    }
    for address in touched {
        let Some(raw) = tx.sheet(&address.sheet_id)?.metadata.get("dataValidations") else {
            continue;
        };
        let rules = raw
            .as_array()
            .ok_or_else(|| invalid("Data validations must be array"))?;
        let mut selected = None;
        for rule in rules {
            let ranges = rule["ranges"]
                .as_array()
                .ok_or_else(|| invalid("Validation ranges must be array"))?;
            for item in ranges {
                if range(tx, &address.sheet_id, item)?.contains(address) {
                    selected = Some(rule);
                    break;
                }
            }
            if selected.is_some() {
                break;
            }
        }
        let Some(rule) = selected else { continue };
        let rid = text(rule, "id")?;
        let decision = p
            .get("writeAuthority")
            .and_then(|v| v.get("validationDecision"));
        if let Some(declared) = decision
            .and_then(|d| d.get("ruleId"))
            .filter(|v| !v.is_null())
        {
            if declared.as_str() != Some(rid) {
                return Err(KernelError::new(
                    "CELL_ENTRY_VALIDATION_STALE",
                    "Validation rule changed before commit",
                )
                .at(rid));
            }
        }
        let value = match tx.read_cell(address)? {
            Some(c) => {
                if let Some(f) = c.formula {
                    evaluate(tx, address, &f)?.scalar()
                } else {
                    c.value
                }
            }
            None => Scalar::Null,
        };
        if valid(tx, address, &value, rule)? {
            continue;
        }
        match rule
            .get("alertStyle")
            .and_then(Value::as_str)
            .unwrap_or("stop")
        {
            "stop" => {
                return Err(KernelError::new(
                    "CELL_ENTRY_VALIDATION_FAILED",
                    "Cell value does not satisfy its validation rule",
                )
                .at(rid));
            }
            "warning" | "information" => {
                if decision
                    .and_then(|d| d.get("status"))
                    .and_then(Value::as_str)
                    != Some("confirmed")
                {
                    return Err(KernelError::new(
                        "CELL_ENTRY_CONFIRMATION_REQUIRED",
                        "Validation warning requires explicit confirmation",
                    )
                    .at(rid)
                    .recover("confirm-validation-warning"));
                }
            }
            _ => return Err(invalid("Invalid validation alert style")),
        }
    }
    Ok(())
}
fn evaluate(tx: &Transaction, address: &CellAddress, formula: &str) -> KernelResult<FormulaValue> {
    let mut runtime = FormulaRuntime::new(&address.sheet_id);
    for sheet in &tx.manifest.sheets {
        runtime.register_sheet(&sheet.name, &sheet.sheet_id)?;
    }
    runtime.context.date1904 = tx
        .manifest
        .metadata
        .get("date1904")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    if let Some(raw) = tx.manifest.metadata.get("definedNameModels") {
        for name in raw
            .as_array()
            .ok_or_else(|| invalid("Defined names must be array"))?
        {
            let scope = match name.get("scope").and_then(Value::as_str).unwrap_or("workbook") {
                "workbook" => DefinedNameScope::Workbook,
                "sheet" => DefinedNameScope::Sheet(text(name, "sheetId")?.into()),
                _ => return Err(invalid("Defined-name scope must be workbook or sheet")),
            };
            runtime.define_name(text(name, "name")?, text(name, "formula")?, scope, address)?;
        }
    }
    for sheet in &tx.manifest.sheets {
        if let Some(raw) = sheet.metadata.get("sheetTables") {
            for table in raw
                .as_array()
                .ok_or_else(|| invalid("Sheet tables must be array"))?
            {
                let columns = table
                    .get("columns")
                    .and_then(Value::as_array)
                    .ok_or_else(|| invalid("Sheet table columns must be array"))?
                    .iter()
                    .map(|column| text(column, "name").map(str::to_owned))
                    .collect::<KernelResult<Vec<_>>>()?;
                runtime.define_table(FormulaTable {
                    name: text(table, "name")?.into(),
                    range: range(tx, &sheet.sheet_id, &table["range"])?,
                    has_header_row: table.get("hasHeaderRow").and_then(Value::as_bool).ok_or_else(|| invalid("Sheet table hasHeaderRow is required"))?,
                    has_total_row: table.get("hasTotalRow").and_then(Value::as_bool).ok_or_else(|| invalid("Sheet table hasTotalRow is required"))?,
                    columns,
                })?;
            }
        }
    }
    runtime.evaluate(
        &if formula.starts_with('=') {
            formula.into()
        } else {
            format!("={formula}")
        },
        address,
        tx,
    )
}
fn as_text(v: &Scalar) -> String {
    match v {
        Scalar::Null => String::new(),
        Scalar::Boolean(v) => {
            if *v {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        Scalar::Text(v) => v.clone(),
        Scalar::Number(v) => v.to_string(),
        Scalar::Error(e) => e.code.clone(),
    }
}
fn bound(tx: &Transaction, a: &CellAddress, rule: &Value, key: &str) -> KernelResult<f64> {
    let v = rule
        .get(key)
        .ok_or_else(|| invalid(format!("Validation {key} required")))?;
    if let Some(n) = v.as_f64() {
        return Ok(n);
    }
    let s = v
        .as_str()
        .ok_or_else(|| invalid("Validation bound must be number or formula"))?;
    if let Ok(n) = s.parse::<f64>() {
        if n.is_finite() {
            return Ok(n);
        }
    }
    match evaluate(tx, a, s)?.scalar() {
        Scalar::Number(n) => Ok(n),
        Scalar::Error(e) => {
            Err(KernelError::new("DATA_VALIDATION_EVALUATION_FAILED", e.message).at(e.code))
        }
        _ => Err(invalid("Validation bound is not numeric")),
    }
}
fn valid(tx: &Transaction, a: &CellAddress, v: &Scalar, rule: &Value) -> KernelResult<bool> {
    if matches!(v, Scalar::Null) || matches!(v,Scalar::Text(s) if s.is_empty()) {
        return Ok(rule
            .get("allowBlank")
            .and_then(Value::as_bool)
            .unwrap_or(true));
    }
    let kind = text(rule, "type")?.to_ascii_lowercase();
    if kind == "checkbox" {
        return Ok(matches!(v, Scalar::Boolean(_))
            || matches!(v,Scalar::Text(s) if s.eq_ignore_ascii_case("true")||s.eq_ignore_ascii_case("false")));
    }
    if kind == "custom" {
        return match evaluate(tx, a, text(rule, "formula1")?)?.scalar() {
            Scalar::Boolean(v) => Ok(v),
            Scalar::Number(n) => Ok(n != 0.),
            Scalar::Error(e) => {
                Err(KernelError::new("DATA_VALIDATION_EVALUATION_FAILED", e.message).at(e.code))
            }
            _ => Ok(false),
        };
    }
    if kind == "list" {
        let target = as_text(v);
        if let Some(source) = rule.get("listSource") {
            match text(source, "kind")? {
                "values" => {
                    for value in source["values"]
                        .as_array()
                        .ok_or_else(|| invalid("Validation values must be array"))?
                    {
                        let scalar: Scalar = serde_json::from_value(value.clone())
                            .map_err(|e| invalid(e.to_string()))?;
                        if as_text(&scalar).eq_ignore_ascii_case(&target) {
                            return Ok(true);
                        }
                    }
                    return Ok(false);
                }
                "range" => {
                    let r: RangeRef = serde_json::from_value(source["range"].clone())
                        .map_err(|e| invalid(e.to_string()))?;
                    range(tx, &r.sheet_id, &source["range"])?;
                    for (address, c) in tx.cells_in_range(&r)? {
                        let scalar = if let Some(f) = c.formula {
                            evaluate(tx, &address, &f)?.scalar()
                        } else {
                            c.value
                        };
                        if as_text(&scalar).eq_ignore_ascii_case(&target) {
                            return Ok(true);
                        }
                    }
                    return Ok(false);
                }
                "formula" => {
                    return Ok(evaluate(tx, a, text(source, "formula")?)?
                        .matrix()
                        .iter()
                        .flatten()
                        .any(|v| as_text(v).eq_ignore_ascii_case(&target)));
                }
                _ => return Err(invalid("Unknown validation list source")),
            }
        }
        let formula = text(rule, "formula1")?;
        let literal = formula.trim_start_matches('=');
        if formula.starts_with('=') && !literal.starts_with('"') {
            return Ok(evaluate(tx, a, formula)?
                .matrix()
                .iter()
                .flatten()
                .any(|v| as_text(v).eq_ignore_ascii_case(&target)));
        }
        return Ok(literal
            .trim_matches('"')
            .split(',')
            .any(|s| s.trim().eq_ignore_ascii_case(&target)));
    }
    let n = if kind == "textlength" {
        as_text(v).encode_utf16().count() as f64
    } else {
        match v {
            Scalar::Number(n) => *n,
            _ => return Ok(false),
        }
    };
    match kind.as_str() {
        "whole" if n.fract() != 0. => return Ok(false),
        "time" if !(0.0..1.0).contains(&n) => return Ok(false),
        "whole" | "decimal" | "date" | "time" | "textlength" => {}
        _ => return Err(invalid("Unknown validation type")),
    }
    let first = bound(tx, a, rule, "formula1")?;
    Ok(
        match rule
            .get("operator")
            .and_then(Value::as_str)
            .unwrap_or("between")
        {
            "greaterThan" => n > first,
            "greaterThanOrEqual" => n >= first,
            "lessThan" => n < first,
            "lessThanOrEqual" => n <= first,
            "equal" => n == first,
            "notEqual" => n != first,
            "between" => n >= first && n <= bound(tx, a, rule, "formula2")?,
            "notBetween" => n < first || n > bound(tx, a, rule, "formula2")?,
            _ => return Err(invalid("Unknown validation operator")),
        },
    )
}
