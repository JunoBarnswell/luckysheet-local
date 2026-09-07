//! Canonical array, lookup, and matrix functions.
//!
//! This module deliberately keeps shape and lookup semantics in one place.  A
//! malformed argument is an Excel value error; only an impossible host
//! failure is returned as `Err` by the caller.
use crate::{FormulaValue, num, text, truth};
use kernel_core::{KernelResult, Scalar};
type ExcelResult<T> = Result<T, FormulaValue>;

pub const FUNCTIONS: &[&str] = &[
    "FILTER",
    "UNIQUE",
    "SORT",
    "SEQUENCE",
    "HSTACK",
    "VSTACK",
    "TAKE",
    "DROP",
    "SORTBY",
    "XMATCH",
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "XLOOKUP",
    "CHOOSE",
    "COLUMNS",
    "ROWS",
    "TRANSPOSE",
    "GROUPBY",
    "PIVOTBY",
];

fn err(code: &str, msg: impl Into<String>) -> FormulaValue {
    FormulaValue::error(code, msg)
}
fn scalar(v: &FormulaValue) -> Scalar {
    v.scalar()
}
fn m(v: &FormulaValue) -> Vec<Vec<Scalar>> {
    match v {
        FormulaValue::Scalar(s) => vec![vec![s.clone()]],
        FormulaValue::Array(a) => a.clone(),
    }
}
fn flat(v: &FormulaValue) -> Vec<Scalar> {
    m(v).into_iter().flatten().collect()
}
fn boolish(s: &Scalar) -> Result<bool, FormulaValue> {
    truth(s).map_err(|e| err(&e.code, e.message))
}
fn numv(v: &FormulaValue) -> Result<f64, FormulaValue> {
    num(&scalar(v)).map_err(|e| err(&e.code, e.message))
}
fn indexv(v: &FormulaValue) -> Result<usize, FormulaValue> {
    let n = numv(v)?;
    if !n.is_finite() || n.fract() != 0.0 {
        return Err(err("#VALUE!", "Expected integer"));
    }
    if n < 0.0 {
        return Err(err("#VALUE!", "Expected non-negative integer"));
    }
    Ok(n as usize)
}
fn cmp(a: &Scalar, b: &Scalar) -> Option<std::cmp::Ordering> {
    match (a, b) {
        (Scalar::Number(x), Scalar::Number(y)) => x.partial_cmp(y),
        _ => Some(
            text(a)
                .to_ascii_lowercase()
                .cmp(&text(b).to_ascii_lowercase()),
        ),
    }
}
fn wildcard(pattern: &str, value: &str) -> bool {
    let p: Vec<char> = pattern.chars().collect();
    let v: Vec<char> = value.chars().collect();
    let mut dp = vec![false; v.len() + 1];
    dp[0] = true;
    let mut i = 0;
    while i < p.len() {
        let c = p[i];
        let mut next = vec![false; v.len() + 1];
        if c == '*' {
            next.clone_from(&dp);
            for j in 1..=v.len() {
                next[j] |= next[j - 1];
            }
        } else if c == '~' && i + 1 < p.len() {
            for j in 1..=v.len() {
                next[j] = dp[j - 1] && p[i + 1].eq_ignore_ascii_case(&v[j - 1]);
            }
            i += 1;
        } else {
            for j in 1..=v.len() {
                next[j] = dp[j - 1] && (c == '?' || c.eq_ignore_ascii_case(&v[j - 1]));
            }
        }
        dp = next;
        i += 1;
    }
    dp[v.len()]
}
fn lookup(value: &Scalar, values: &[Scalar], mode: i32, search: i32) -> Option<usize> {
    if search.abs() == 2 {
        let ascending = search == 2;
        let (mut lo, mut hi) = (0usize, values.len());
        while lo < hi {
            let mid = lo + (hi - lo) / 2;
            let order = cmp(&values[mid], value)?;
            let before = if ascending {
                order.is_lt()
            } else {
                order.is_gt()
            };
            if before { lo = mid + 1 } else { hi = mid }
        }
        if lo < values.len() && cmp(&values[lo], value)?.is_eq() {
            return Some(lo);
        }
        let candidate = match (ascending, mode) {
            (true, -1) | (false, 1) => lo.checked_sub(1),
            (true, 1) | (false, -1) if lo < values.len() => Some(lo),
            _ => None,
        };
        return candidate;
    }
    let indices: Box<dyn Iterator<Item = usize>> = if search == -1 {
        Box::new((0..values.len()).rev())
    } else {
        Box::new(0..values.len())
    };
    let mut best: Option<usize> = None;
    for i in indices {
        let order = cmp(&values[i], value)?;
        if mode == 2 {
            if let (Scalar::Text(pattern), Scalar::Text(candidate)) = (value, &values[i]) {
                if wildcard(pattern, candidate) {
                    return Some(i);
                }
            } else if order.is_eq() {
                return Some(i);
            }
            continue;
        }
        if order.is_eq() {
            return Some(i);
        }
        if (mode == -1 && order.is_lt()) || (mode == 1 && order.is_gt()) {
            let better = match best {
                None => true,
                Some(previous) => {
                    let difference = cmp(&values[i], &values[previous])?;
                    if mode == -1 {
                        difference.is_gt()
                    } else {
                        difference.is_lt()
                    }
                }
            };
            if better {
                best = Some(i)
            }
        }
    }
    best
}
fn modes(args: &[FormulaValue], mo: usize, so: usize) -> Result<(i32, i32), FormulaValue> {
    let a = if args.len() > mo {
        numv(&args[mo])? as i32
    } else {
        0
    };
    let b = if args.len() > so {
        numv(&args[so])? as i32
    } else {
        1
    };
    if ![0, -1, 1, 2].contains(&a) || ![1, -1, 2, -2].contains(&b) || (a == 2 && b.abs() == 2) {
        Err(err("#VALUE!", "Invalid lookup mode"))
    } else {
        Ok((a, b))
    }
}

/// Dispatch with the host-provided dynamic-array resource budget.  The
/// evaluator owns the budget; this module never truncates an array.
pub fn call_with_limit(
    name: &str,
    args: &[FormulaValue],
    max_cells: u64,
) -> Option<KernelResult<FormulaValue>> {
    let n = name.to_ascii_uppercase();
    if !FUNCTIONS.contains(&n.as_str()) {
        return None;
    }
    if let Err(budget) = preflight_budget(&n, args, max_cells) {
        return Some(Err(budget));
    }
    let r = match n.as_str() {
        "FILTER" => filter(args),
        "UNIQUE" => unique(args),
        "SORT" => sort(args),
        "SEQUENCE" => sequence(args),
        "HSTACK" => stack(args, true),
        "VSTACK" => stack(args, false),
        "TAKE" => take_drop(args, false),
        "DROP" => take_drop(args, true),
        "SORTBY" => sortby(args),
        "XMATCH" => xmatch(args),
        "VLOOKUP" => vlookup(args, false),
        "HLOOKUP" => vlookup(args, true),
        "INDEX" => index(args),
        "MATCH" => xlookup_match(args),
        "XLOOKUP" => xlookup(args),
        "CHOOSE" => choose(args),
        "COLUMNS" => Ok(FormulaValue::Scalar(Scalar::Number(
            m(args.first().unwrap_or(&FormulaValue::Scalar(Scalar::Null)))
                .first()
                .map(|r| r.len())
                .unwrap_or(0) as f64,
        ))),
        "ROWS" => Ok(FormulaValue::Scalar(Scalar::Number(
            m(args.first().unwrap_or(&FormulaValue::Scalar(Scalar::Null))).len() as f64,
        ))),
        "TRANSPOSE" => transpose(args),
        "GROUPBY" => groupby(args, false),
        "PIVOTBY" => groupby(args, true),
        _ => unreachable!(),
    };
    Some(match r {
        Ok(value) => {
            if let FormulaValue::Array(ref matrix) = value {
                let count = matrix
                    .iter()
                    .try_fold(0u64, |n, row| n.checked_add(row.len() as u64));
                if count.map(|n| n > max_cells).unwrap_or(true) {
                    return Some(Err(kernel_core::KernelError::new(
                        "RESOURCE_BUDGET_EXCEEDED",
                        "Dynamic array exceeds calculation resource budget",
                    )));
                }
            }
            Ok(value)
        }
        Err(value) => Ok(value),
    })
}

fn preflight_budget(name: &str, args: &[FormulaValue], limit: u64) -> KernelResult<()> {
    let cells = |rows: u64, cols: u64| -> KernelResult<()> {
        if rows.checked_mul(cols).map(|n| n > limit).unwrap_or(true) {
            Err(kernel_core::KernelError::new(
                "RESOURCE_BUDGET_EXCEEDED",
                "Dynamic array exceeds calculation resource budget",
            ))
        } else {
            Ok(())
        }
    };
    match name {
        "SEQUENCE" => {
            if !args.is_empty() {
                if let (Ok(r), Ok(c)) = (numv(&args[0]), args.get(1).map(numv).unwrap_or(Ok(1.))) {
                    if r.is_finite() && c.is_finite() && r >= 0. && c >= 0. {
                        cells(r as u64, c as u64)?;
                    }
                }
            }
        }
        "HSTACK" => {
            let ms: Vec<_> = args.iter().map(m).collect();
            let rows = ms.iter().map(|x| x.len()).max().unwrap_or(0) as u64;
            let cols = ms
                .iter()
                .map(|x| x.iter().map(Vec::len).max().unwrap_or(0) as u64)
                .sum();
            cells(rows, cols)?;
        }
        "VSTACK" => {
            let ms: Vec<_> = args.iter().map(m).collect();
            let rows = ms.iter().map(|x| x.len() as u64).sum();
            let cols = ms
                .iter()
                .map(|x| x.iter().map(Vec::len).max().unwrap_or(0) as u64)
                .max()
                .unwrap_or(0);
            cells(rows, cols)?;
        }
        "FILTER" => {
            let x = m(args.first().unwrap_or(&FormulaValue::Scalar(Scalar::Null)));
            let rows = x.len() as u64;
            let cols = x.iter().map(Vec::len).max().unwrap_or(0) as u64;
            cells(rows, cols)?;
        }
        "GROUPBY" | "PIVOTBY" => {
            let fields = m(args.first().unwrap_or(&FormulaValue::Scalar(Scalar::Null)));
            let value_index = if name == "PIVOTBY" { 2 } else { 1 };
            let values = m(args
                .get(value_index)
                .unwrap_or(&FormulaValue::Scalar(Scalar::Null)));
            let row_groups = fields
                .iter()
                .map(|row| group_key(row))
                .collect::<std::collections::BTreeSet<_>>()
                .len() as u64;
            let field_width = fields.first().map_or(0, Vec::len) as u64;
            let value_width = values.first().map_or(0, Vec::len) as u64;
            if name == "PIVOTBY" && args.len() > 1 {
                let columns = m(&args[1]);
                let col_groups = columns
                    .iter()
                    .map(|row| group_key(row))
                    .collect::<std::collections::BTreeSet<_>>()
                    .len() as u64;
                cells(
                    row_groups + columns.first().map_or(0, Vec::len) as u64,
                    field_width + col_groups * value_width,
                )?;
            } else {
                cells(row_groups, field_width + value_width)?;
            }
        }
        _ => {}
    }
    Ok(())
}

fn require(args: &[FormulaValue], n: usize, f: &str) -> Result<(), FormulaValue> {
    if args.len() < n {
        Err(err("#VALUE!", format!("{f} requires {n} arguments")))
    } else {
        Ok(())
    }
}
fn filter(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    require(a, 2, "FILTER")?;
    let x = m(&a[0]);
    let inc = m(&a[1]);
    let width = x.iter().map(Vec::len).max().unwrap_or(0);
    let row_mode = inc.len() == x.len() && inc.iter().all(|r| r.len() == 1);
    let col_mode = inc.len() == 1 && inc.first().map(Vec::len).unwrap_or(0) == width;
    if !row_mode && !col_mode {
        return Ok(err("#VALUE!", "FILTER include shape mismatch"));
    }
    let mut out = Vec::new();
    if col_mode {
        let selected: Vec<bool> = inc[0].iter().map(boolish).collect::<Result<_, _>>()?;
        for row in &x {
            out.push(
                row.iter()
                    .enumerate()
                    .filter_map(|(j, v)| {
                        selected
                            .get(j)
                            .copied()
                            .unwrap_or(false)
                            .then_some(v.clone())
                    })
                    .collect(),
            );
        }
    } else {
        for (i, row) in x.iter().enumerate() {
            if boolish(&inc[i][0])? {
                out.push(row.clone())
            }
        }
    }
    if out.is_empty() {
        Ok(if a.len() > 2 {
            FormulaValue::Array(m(&a[2]))
        } else {
            err("#CALC!", "FILTER returned no results")
        })
    } else {
        Ok(FormulaValue::Array(out))
    }
}
fn unique(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    require(a, 1, "UNIQUE")?;
    let x = m(&a[0]);
    let by = if let Some(v) = a.get(1) {
        boolish(&scalar(v))?
    } else {
        false
    };
    let once = if let Some(v) = a.get(2) {
        boolish(&scalar(v))?
    } else {
        false
    };
    if by {
        return Ok(FormulaValue::Array(unique_cols(&x, once)));
    }
    let mut out = Vec::new();
    for r in &x {
        let c = x.iter().filter(|q| *q == r).count();
        if (!once || c == 1) && !out.contains(r) {
            out.push(r.clone())
        }
    }
    Ok(FormulaValue::Array(if out.is_empty() {
        vec![vec![]]
    } else {
        out
    }))
}
fn unique_cols(x: &[Vec<Scalar>], once: bool) -> Vec<Vec<Scalar>> {
    let w = x.iter().map(Vec::len).max().unwrap_or(0);
    let mut cols = Vec::new();
    for c in 0..w {
        let z = x
            .iter()
            .map(|r| r.get(c).cloned().unwrap_or_default())
            .collect::<Vec<_>>();
        let count = (0..w)
            .filter(|&j| {
                x.iter()
                    .map(|r| r.get(j).cloned().unwrap_or_default())
                    .collect::<Vec<_>>()
                    == z
            })
            .count();
        if (!once || count == 1) && !cols.contains(&z) {
            cols.push(z)
        }
    }
    let h = x.len();
    (0..h)
        .map(|r| cols.iter().map(|c| c[r].clone()).collect())
        .collect()
}
fn sort(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    require(a, 1, "SORT")?;
    let mut x = m(&a[0]);
    let idx = if a.len() > 1 {
        indexv(&a[1])?.saturating_sub(1)
    } else {
        0
    };
    let rev = if a.len() > 2 {
        numv(&a[2])? < 0.0
    } else {
        false
    };
    let by = if a.len() > 3 {
        boolish(&scalar(&a[3]))?
    } else {
        false
    };
    if by {
        if idx >= x.iter().map(Vec::len).max().unwrap_or(0) {
            return Ok(err("#VALUE!", "SORT index out of bounds"));
        }
        let mut ix: Vec<usize> = (0..x.iter().map(Vec::len).max().unwrap_or(0)).collect();
        ix.sort_by(|i, j| {
            cmp(
                x[0].get(*i).unwrap_or(&Scalar::Null),
                x[0].get(*j).unwrap_or(&Scalar::Null),
            )
            .unwrap_or(std::cmp::Ordering::Equal)
        });
        if rev {
            ix.reverse()
        }
        Ok(FormulaValue::Array(
            x.iter()
                .map(|r| {
                    ix.iter()
                        .map(|i| r.get(*i).cloned().unwrap_or_default())
                        .collect()
                })
                .collect(),
        ))
    } else {
        if idx >= x.iter().map(Vec::len).max().unwrap_or(0) {
            return Ok(err("#VALUE!", "SORT index out of bounds"));
        }
        x.sort_by(|r, s| {
            cmp(
                r.get(idx).unwrap_or(&Scalar::Null),
                s.get(idx).unwrap_or(&Scalar::Null),
            )
            .unwrap_or(std::cmp::Ordering::Equal)
        });
        if rev {
            x.reverse()
        }
        Ok(FormulaValue::Array(x))
    }
}
fn sequence(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    let r = indexv(
        a.first()
            .ok_or_else(|| err("#VALUE!", "SEQUENCE rows required"))?,
    )?;
    let c = if a.len() > 1 { indexv(&a[1])? } else { 1 };
    let st = if a.len() > 2 { numv(&a[2])? } else { 1. };
    let step = if a.len() > 3 { numv(&a[3])? } else { 1. };
    Ok(FormulaValue::Array(
        (0..r)
            .map(|i| {
                (0..c)
                    .map(|j| Scalar::Number(st + (i * c + j) as f64 * step))
                    .collect()
            })
            .collect(),
    ))
}
fn stack(a: &[FormulaValue], horizontal: bool) -> ExcelResult<FormulaValue> {
    if a.is_empty() {
        return Ok(err("#VALUE!", "stack requires arrays"));
    }
    let z = a.iter().map(m).collect::<Vec<_>>();
    if horizontal {
        let h = z.iter().map(Vec::len).max().unwrap_or(0);
        Ok(FormulaValue::Array(
            (0..h)
                .map(|r| {
                    z.iter()
                        .flat_map(|q| q.get(r).cloned().unwrap_or_default())
                        .collect()
                })
                .collect(),
        ))
    } else {
        let w = z
            .iter()
            .flat_map(|q| q.iter().map(Vec::len))
            .max()
            .unwrap_or(0);
        Ok(FormulaValue::Array(
            z.iter()
                .flat_map(|q| {
                    q.iter().map(|r| {
                        let mut x = r.clone();
                        x.resize(w, Scalar::Null);
                        x
                    })
                })
                .collect(),
        ))
    }
}
fn take_drop(a: &[FormulaValue], drop: bool) -> ExcelResult<FormulaValue> {
    require(a, 2, if drop { "DROP" } else { "TAKE" })?;
    let x = m(&a[0]);
    let n = numv(&a[1])? as isize;
    let cc = if a.len() > 2 {
        Some(numv(&a[2])? as isize)
    } else {
        None
    };
    let rows = if drop {
        if n >= 0 {
            x.get(n as usize..).unwrap_or(&[]).to_vec()
        } else {
            x[..x.len().saturating_sub((-n) as usize)].to_vec()
        }
    } else {
        if n >= 0 {
            x[..(n as usize).min(x.len())].to_vec()
        } else {
            x[x.len().saturating_sub((-n) as usize)..].to_vec()
        }
    };
    let rows = if let Some(c) = cc {
        rows.into_iter()
            .map(|r| {
                if drop {
                    if c >= 0 {
                        r.get(c as usize..).unwrap_or(&[]).to_vec()
                    } else {
                        r[..r.len().saturating_sub((-c) as usize)].to_vec()
                    }
                } else {
                    if c >= 0 {
                        r[..(c as usize).min(r.len())].to_vec()
                    } else {
                        r[r.len().saturating_sub((-c) as usize)..].to_vec()
                    }
                }
            })
            .collect()
    } else {
        rows
    };
    Ok(FormulaValue::Array(rows))
}
fn sortby(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    require(a, 2, "SORTBY")?;
    let x = m(&a[0]);
    let key = flat(&a[1]);
    let rev = if let Some(v) = a.get(2) {
        numv(v)? < 0.
    } else {
        false
    };
    if key.len() != x.len() {
        return Ok(err("#VALUE!", "SORTBY sort array size mismatch"));
    }
    let mut order: Vec<usize> = (0..x.len()).collect();
    order.sort_by(|i, j| {
        let o = cmp(&key[*i], &key[*j]).unwrap_or(std::cmp::Ordering::Equal);
        if rev { o.reverse() } else { o }
    });
    Ok(FormulaValue::Array(
        order.into_iter().map(|i| x[i].clone()).collect(),
    ))
}
fn xmatch(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    require(a, 2, "XMATCH")?;
    let (mo, so) = modes(a, 2, 3)?;
    let i = lookup(&scalar(&a[0]), &lookup_vector(&a[1])?, mo, so);
    Ok(
        i.map(|x| FormulaValue::Scalar(Scalar::Number((x + 1) as f64)))
            .unwrap_or_else(|| err("#N/A", "Value not found in XMATCH")),
    )
}
fn xlookup_match(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    require(a, 2, "MATCH")?;
    let mo = if a.len() > 2 { numv(&a[2])? as i32 } else { 1 };
    if ![-1, 0, 1].contains(&mo) {
        return Err(err("#N/A", "MATCH match_type must be -1, 0, or 1"));
    }
    let values = lookup_vector(&a[1])?;
    Ok(
        lookup(&scalar(&a[0]), &values, if mo == 0 { 2 } else { -mo }, 1)
            .map(|x| FormulaValue::Scalar(Scalar::Number((x + 1) as f64)))
            .unwrap_or_else(|| err("#N/A", "Value not found in MATCH")),
    )
}
fn lookup_vector(value: &FormulaValue) -> ExcelResult<Vec<Scalar>> {
    let rows = matrix(value, "Lookup vector")?;
    if rows.len() == 1 {
        Ok(rows[0].clone())
    } else if rows[0].len() == 1 {
        Ok(rows.into_iter().map(|row| row[0].clone()).collect())
    } else {
        Err(err("#VALUE!", "Lookup array must be one row or one column"))
    }
}
pub(crate) fn xlookup_outcome(a: &[FormulaValue]) -> ExcelResult<Option<FormulaValue>> {
    require(a, 3, "XLOOKUP")?;
    let lm = matrix(&a[1], "XLOOKUP lookup_array")?;
    let rm = matrix(&a[2], "XLOOKUP return_array")?;
    let horizontal = lm.len() == 1;
    if (!horizontal && lm[0].len() != 1)
        || (horizontal && rm[0].len() != lm[0].len())
        || (!horizontal && rm.len() != lm.len())
    {
        return Err(err(
            "#VALUE!",
            "XLOOKUP lookup and return array shapes differ",
        ));
    }
    let value = scalar(&a[0]);
    if matches!(value, Scalar::Error(_)) {
        return Err(FormulaValue::Scalar(value));
    }
    let lookup_values = lookup_vector(&a[1])?;
    let (mode, search) = modes(a, 4, 5)?;
    let Some(index) = lookup(&value, &lookup_values, mode, search) else {
        return Ok(None);
    };
    Ok(Some(if horizontal {
        if rm.len() == 1 {
            FormulaValue::Scalar(rm[0][index].clone())
        } else {
            FormulaValue::Array(rm.into_iter().map(|row| vec![row[index].clone()]).collect())
        }
    } else if rm[index].len() == 1 {
        FormulaValue::Scalar(rm[index][0].clone())
    } else {
        FormulaValue::Array(vec![rm[index].clone()])
    }))
}
fn xlookup(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    Ok(xlookup_outcome(a)?.unwrap_or_else(|| {
        a.get(3)
            .cloned()
            .unwrap_or_else(|| err("#N/A", "Value not found in XLOOKUP"))
    }))
}
fn vlookup(a: &[FormulaValue], horizontal: bool) -> ExcelResult<FormulaValue> {
    require(a, 3, if horizontal { "HLOOKUP" } else { "VLOOKUP" })?;
    let t = m(&a[1]);
    let n = indexv(&a[2])?;
    if n == 0 {
        return Err(err("#VALUE!", "Lookup column/row index must be positive"));
    }
    let exact = if let Some(v) = a.get(3) {
        !boolish(&scalar(v))?
    } else {
        false
    };
    let (vec, idx) = if horizontal {
        (t.first().cloned().unwrap_or_default(), n.saturating_sub(1))
    } else {
        (
            t.iter()
                .map(|r| r.first().cloned().unwrap_or_default())
                .collect(),
            n.saturating_sub(1),
        )
    };
    if idx
        >= if horizontal {
            t.len()
        } else {
            t.first().map(Vec::len).unwrap_or(0)
        }
    {
        return Ok(err("#REF!", "Lookup index out of bounds"));
    }
    let Some(i) = lookup(&scalar(&a[0]), &vec, if exact { 0 } else { -1 }, 1) else {
        return Ok(err("#N/A", "Value not found"));
    };
    Ok(FormulaValue::Scalar(if horizontal {
        t[idx].get(i).cloned().unwrap_or_default()
    } else {
        t[i].get(idx).cloned().unwrap_or_default()
    }))
}
fn index(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    require(a, 1, "INDEX")?;
    let t = m(&a[0]);
    let r = if a.len() > 1 { indexv(&a[1])? } else { 1 };
    let c = if a.len() > 2 { indexv(&a[2])? } else { 1 };
    let width = t.first().map_or(0, Vec::len);
    if r > t.len() || c > width {
        return Err(err("#REF!", "INDEX out of bounds"));
    }
    if r == 0 && c == 0 {
        return Ok(FormulaValue::Array(t));
    }
    if r == 0 {
        return Ok(FormulaValue::Array(
            t.iter()
                .map(|x| vec![x.get(c.saturating_sub(1)).cloned().unwrap_or_default()])
                .collect(),
        ));
    }
    if c == 0 {
        return Ok(FormulaValue::Array(vec![
            t.get(r - 1).cloned().unwrap_or_default(),
        ]));
    }
    Ok(t.get(r - 1)
        .and_then(|x| x.get(c - 1))
        .cloned()
        .map(|x| FormulaValue::Scalar(x))
        .unwrap_or_else(|| err("#REF!", "INDEX out of bounds")))
}
fn choose(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    require(a, 2, "CHOOSE")?;
    let i = indexv(&a[0])?;
    Ok(a.get(i)
        .cloned()
        .unwrap_or_else(|| err("#VALUE!", "CHOOSE index out of bounds")))
}
fn transpose(a: &[FormulaValue]) -> ExcelResult<FormulaValue> {
    let x = m(a.first().unwrap_or(&FormulaValue::Scalar(Scalar::Null)));
    let w = x.iter().map(Vec::len).max().unwrap_or(0);
    Ok(FormulaValue::Array(
        (0..w)
            .map(|c| {
                x.iter()
                    .map(|r| r.get(c).cloned().unwrap_or_default())
                    .collect()
            })
            .collect(),
    ))
}

fn matrix(v: &FormulaValue, label: &str) -> Result<Vec<Vec<Scalar>>, FormulaValue> {
    let x = m(v);
    if x.is_empty() || x.iter().any(|r| r.is_empty() || r.len() != x[0].len()) {
        return Err(err(
            "#VALUE!",
            format!("{label} must be rectangular and non-empty"),
        ));
    }
    Ok(x)
}
#[derive(Clone)]
struct GroupStats {
    count: u64,
    sum: f64,
    min: f64,
    max: f64,
}
impl GroupStats {
    fn new() -> Self {
        Self {
            count: 0,
            sum: 0.,
            min: f64::INFINITY,
            max: f64::NEG_INFINITY,
        }
    }
    fn add(&mut self, value: &Scalar) -> ExcelResult<()> {
        match value {
            Scalar::Number(n) => {
                self.count += 1;
                self.sum += n;
                self.min = self.min.min(*n);
                self.max = self.max.max(*n);
            }
            Scalar::Error(_) => return Err(FormulaValue::Scalar(value.clone())),
            _ => {}
        }
        Ok(())
    }
    fn value(&self, aggregation: &str) -> Scalar {
        match aggregation {
            "COUNT" => Scalar::Number(self.count as f64),
            "AVERAGE" if self.count == 0 => Scalar::error("#DIV/0!", "No numeric observations"),
            "AVERAGE" => Scalar::Number(self.sum / self.count as f64),
            "MIN" => Scalar::Number(if self.count == 0 { 0. } else { self.min }),
            "MAX" => Scalar::Number(if self.count == 0 { 0. } else { self.max }),
            _ => Scalar::Number(self.sum),
        }
    }
}
fn group_key(row: &[Scalar]) -> String {
    let normalized: Vec<_> = row
        .iter()
        .map(|v| {
            if let Scalar::Text(text) = v {
                Scalar::Text(text.to_lowercase())
            } else {
                v.clone()
            }
        })
        .collect();
    serde_json::to_string(&normalized).expect("validated scalar group key serializes")
}
fn groupby(a: &[FormulaValue], pivot: bool) -> ExcelResult<FormulaValue> {
    let required = if pivot { 4 } else { 3 };
    if a.len() != required {
        return Err(err(
            "#VALUE!",
            "GROUPBY/PIVOTBY require their canonical field, value, and aggregation arguments",
        ));
    }
    let fields = matrix(&a[0], "Row fields")?;
    let columns = if pivot {
        Some(matrix(&a[1], "Column fields")?)
    } else {
        None
    };
    let values = matrix(&a[if pivot { 2 } else { 1 }], "Values")?;
    if fields.len() != values.len() || columns.as_ref().is_some_and(|c| c.len() != values.len()) {
        return Err(err(
            "#VALUE!",
            "Grouping fields and values must have equal row counts",
        ));
    }
    let aggregation = text(&a[required - 1].scalar()).to_uppercase();
    if !["SUM", "COUNT", "AVERAGE", "MIN", "MAX"].contains(&aggregation.as_str()) {
        return Err(err("#VALUE!", "Unsupported grouping aggregation"));
    }
    let mut row_ids = std::collections::BTreeMap::new();
    let mut row_fields = Vec::new();
    let mut col_ids = std::collections::BTreeMap::new();
    let mut col_fields = Vec::new();
    let mut groups: std::collections::BTreeMap<(usize, usize), Vec<GroupStats>> =
        std::collections::BTreeMap::new();
    for i in 0..values.len() {
        let next_row = row_fields.len();
        let row = *row_ids.entry(group_key(&fields[i])).or_insert_with(|| {
            row_fields.push(fields[i].clone());
            next_row
        });
        let col = if let Some(columns) = &columns {
            let next_col = col_fields.len();
            *col_ids.entry(group_key(&columns[i])).or_insert_with(|| {
                col_fields.push(columns[i].clone());
                next_col
            })
        } else {
            0
        };
        let stats = groups
            .entry((row, col))
            .or_insert_with(|| vec![GroupStats::new(); values[0].len()]);
        for (stat, value) in stats.iter_mut().zip(&values[i]) {
            stat.add(value)?;
        }
    }
    let mut output = Vec::new();
    if let Some(columns) = &columns {
        for level in 0..columns[0].len() {
            let mut header = vec![Scalar::Null; fields[0].len()];
            for col in &col_fields {
                header.extend(std::iter::repeat_n(col[level].clone(), values[0].len()));
            }
            output.push(header);
        }
    }
    for (row, fields) in row_fields.into_iter().enumerate() {
        let mut output_row = fields;
        for col in 0..if pivot { col_fields.len() } else { 1 } {
            if let Some(stats) = groups.get(&(row, col)) {
                output_row.extend(stats.iter().map(|s| s.value(&aggregation)));
            } else {
                output_row
                    .extend((0..values[0].len()).map(|_| GroupStats::new().value(&aggregation)));
            }
        }
        output.push(output_row);
    }
    Ok(FormulaValue::Array(output))
}
