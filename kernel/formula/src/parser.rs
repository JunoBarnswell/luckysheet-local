//! Excel formula lexer, parser, and canonical reference resolver.
//!
//! The parser deliberately resolves A1/R1C1 references while parsing.  The
//! resulting AST therefore contains the same `RangeRef` shape used by the
//! kernel dependency and persistence layers.

use kernel_core::{CellAddress, KernelResult, MAX_COLUMNS, MAX_ROWS, RangeRef, Scalar};
use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub enum Expr {
    Scalar(Scalar),
    Reference(RangeRef),
    Name(String),
    Structured(StructuredReference),
    Array(Vec<Vec<Expr>>),
    Unary(String, Box<Expr>),
    Binary(String, Box<Expr>, Box<Expr>),
    Call(String, Vec<Expr>),
    Invoke(Box<Expr>, Vec<Expr>),
    Spill(Box<Expr>),
    Missing,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub struct StructuredReference {
    pub source: String,
    pub table_name: String,
    pub specifier: Option<String>,
    pub column_name: Option<String>,
    pub column_end_name: Option<String>,
    pub this_row: bool,
}

fn err(message: impl Into<String>) -> kernel_core::KernelError {
    kernel_core::KernelError::new("FORMULA_PARSE", message)
}

pub fn parse(source: &str, current: &CellAddress) -> KernelResult<Expr> {
    current.validate()?;
    let normalized = normalize_external_formula(source);
    resolve(&crate::editor::parse_editor(&normalized, current)?, current)
}
pub fn normalize_external_formula(source: &str) -> String {
    let mut output = String::with_capacity(source.len());
    let mut offset = 0;
    let mut in_string = false;
    while offset < source.len() {
        let tail = &source[offset..];
        if in_string && tail.starts_with("\"\"") {
            output.push_str("\"\"");
            offset += 2;
            continue;
        }
        if tail.starts_with('"') {
            in_string = !in_string;
            output.push('"');
            offset += 1;
            continue;
        }
        if !in_string
            && tail
                .get(..6)
                .is_some_and(|prefix| prefix.eq_ignore_ascii_case("_xlfn.") || prefix.eq_ignore_ascii_case("_xlws."))
        {
            offset += 6;
            continue;
        }
        let character = tail.chars().next().expect("offset is within source");
        output.push(character);
        offset += character.len_utf8();
    }
    output
}
fn normalize_function(s: &str) -> String {
    let mut name = s;
    loop {
        let lower = name.to_ascii_lowercase();
        if lower.starts_with("_xlfn.") || lower.starts_with("_xlws.") {
            name = &name[6..];
        } else {
            break;
        }
    }
    name.to_ascii_uppercase()
}
fn resolve(source: &crate::editor::Expr, current: &CellAddress) -> KernelResult<Expr> {
    use crate::editor::Node;
    let child = |e: &crate::editor::Expr| resolve(e, current);
    Ok(match &source.node {
        Node::NumberLiteral { value } => Expr::Scalar(Scalar::Number(*value)),
        Node::StringLiteral { value } => Expr::Scalar(Scalar::Text(value.clone())),
        Node::BooleanLiteral { value } => Expr::Scalar(Scalar::Boolean(*value)),
        Node::ErrorLiteral { code } | Node::InvalidReference { code } => {
            Expr::Scalar(Scalar::error(code, "Formula error literal"))
        }
        Node::NameReference { name } => Expr::Name(name.clone()),
        Node::MissingArgument {} => Expr::Missing,
        Node::CellReference { reference: r } => Expr::Reference(RangeRef {
            sheet_id: r
                .sheet_id
                .clone()
                .unwrap_or_else(|| current.sheet_id.clone()),
            start_row: r.row,
            end_row: r.row,
            start_column: r.column,
            end_column: r.column,
        }),
        Node::RangeReference { start, end } => {
            let Expr::Reference(a) = child(start)? else {
                return Err(err("Invalid range start"));
            };
            let mut range_current = current.clone();
            range_current.sheet_id = a.sheet_id.clone();
            let Expr::Reference(b) = resolve(end, &range_current)? else {
                return Err(err("Invalid range end"));
            };
            if !a.sheet_id.eq_ignore_ascii_case(&b.sheet_id) {
                return Err(kernel_core::KernelError::new(
                    "UNSUPPORTED_FEATURE",
                    "Cross-sheet range endpoints",
                ));
            }
            Expr::Reference(RangeRef {
                sheet_id: a.sheet_id,
                start_row: a.start_row.min(b.start_row),
                end_row: a.end_row.max(b.end_row),
                start_column: a.start_column.min(b.start_column),
                end_column: a.end_column.max(b.end_column),
            })
        }
        Node::WholeColumnReference {
            sheet_id,
            start_column,
            end_column,
            ..
        } => Expr::Reference(RangeRef {
            sheet_id: sheet_id.clone().unwrap_or_else(|| current.sheet_id.clone()),
            start_row: 0,
            end_row: MAX_ROWS - 1,
            start_column: (*start_column).min(*end_column),
            end_column: (*start_column).max(*end_column),
        }),
        Node::WholeRowReference {
            sheet_id,
            start_row,
            end_row,
            ..
        } => Expr::Reference(RangeRef {
            sheet_id: sheet_id.clone().unwrap_or_else(|| current.sheet_id.clone()),
            start_column: 0,
            end_column: MAX_COLUMNS - 1,
            start_row: (*start_row).min(*end_row),
            end_row: (*start_row).max(*end_row),
        }),
        Node::UnaryExpression { operator, operand } => {
            Expr::Unary(operator.clone(), Box::new(child(operand)?))
        }
        Node::BinaryExpression {
            operator,
            left,
            right,
        } => Expr::Binary(
            operator.clone(),
            Box::new(child(left)?),
            Box::new(child(right)?),
        ),
        Node::FunctionCall { name, arguments } => {
            let name = normalize_function(name);
            let arguments = arguments.iter().map(child).collect::<KernelResult<Vec<_>>>()?;
            if name == "SINGLE" && arguments.len() == 1 {
                Expr::Unary("@".into(), Box::new(arguments.into_iter().next().unwrap()))
            } else {
                Expr::Call(name, arguments)
            }
        }
        Node::LambdaInvocation { callee, arguments } => Expr::Invoke(
            Box::new(child(callee)?),
            arguments.iter().map(child).collect::<KernelResult<_>>()?,
        ),
        Node::SpillReference { operand } => Expr::Spill(Box::new(child(operand)?)),
        Node::ArrayLiteral { rows } => Expr::Array(
            rows.iter()
                .map(|row| row.iter().map(child).collect())
                .collect::<KernelResult<_>>()?,
        ),
        Node::TableReference {
            table_name,
            specifier,
            column_name,
            column_end_name,
            this_row,
        } => Expr::Structured(StructuredReference {
            source: crate::editor::format_editor(source),
            table_name: table_name.clone(),
            specifier: specifier.clone(),
            column_name: column_name.clone(),
            column_end_name: column_end_name.clone(),
            this_row: *this_row,
        }),
        Node::ReferenceIntersection { left, right } => {
            let (Expr::Reference(a), Expr::Reference(b)) = (child(left)?, child(right)?) else {
                return Err(kernel_core::KernelError::new(
                    "UNSUPPORTED_FEATURE",
                    "Dynamic reference intersection",
                ));
            };
            if a.sheet_id != b.sheet_id
                || a.start_row > b.end_row
                || b.start_row > a.end_row
                || a.start_column > b.end_column
                || b.start_column > a.end_column
            {
                Expr::Scalar(Scalar::error("#NULL!", "References do not intersect"))
            } else {
                Expr::Reference(RangeRef {
                    sheet_id: a.sheet_id,
                    start_row: a.start_row.max(b.start_row),
                    end_row: a.end_row.min(b.end_row),
                    start_column: a.start_column.max(b.start_column),
                    end_column: a.end_column.min(b.end_column),
                })
            }
        }
        Node::ReferenceUnion { references } => {
            let mut resolved = references.iter().map(child);
            let first = resolved
                .next()
                .ok_or_else(|| err("Empty reference union"))??;
            resolved.try_fold(first, |left, right| {
                Ok(Expr::Binary(",".into(), Box::new(left), Box::new(right?)))
            })?
        }
    })
}

/// Formats the resolved AST into a valid formula expression. Absolute A1
/// markers cannot be reconstructed after resolution, by design; callers that
/// need source-preserving editing should retain the source-side syntax model.
pub fn format(expr: &Expr) -> String {
    match expr {
        Expr::Scalar(Scalar::Null) => String::new(),
        Expr::Scalar(Scalar::Boolean(v)) => {
            if *v {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        Expr::Scalar(Scalar::Number(v)) => v.to_string(),
        Expr::Scalar(Scalar::Text(v)) => format!("\"{}\"", v.replace('"', "\"\"")),
        Expr::Scalar(Scalar::Error(e)) => e.code.clone(),
        Expr::Reference(r) => format_range(r),
        Expr::Name(v) => v.clone(),
        Expr::Structured(v) => v.source.clone(),
        Expr::Array(rows) => format!(
            "{{{}}}",
            rows.iter()
                .map(|row| row.iter().map(format).collect::<Vec<_>>().join(","))
                .collect::<Vec<_>>()
                .join(";")
        ),
        Expr::Unary(op, e) => {
            if op == "%" {
                format!("{}%", format(e))
            } else {
                format!("{}({})", op, format(e))
            }
        }
        Expr::Binary(op, l, r) => format!("({}{}{})", format(l), op, format(r)),
        Expr::Call(name, args) => format!(
            "{}({})",
            name,
            args.iter().map(format).collect::<Vec<_>>().join(",")
        ),
        Expr::Invoke(callee, args) => format!(
            "{}({})",
            format(callee),
            args.iter().map(format).collect::<Vec<_>>().join(",")
        ),
        Expr::Spill(e) => format!("{}#", format(e)),
        Expr::Missing => String::new(),
    }
}

fn format_range(r: &RangeRef) -> String {
    let sheet = if r
        .sheet_id
        .chars()
        .all(|c| c.is_ascii_alphanumeric() || c == '_')
    {
        r.sheet_id.clone()
    } else {
        format!("'{}'", r.sheet_id.replace('\'', "''"))
    };
    let a = format_a1(r.start_row, r.start_column);
    let b = format_a1(r.end_row, r.end_column);
    if a == b {
        format!("{}!{}", sheet, a)
    } else {
        format!("{}!{}:{}", sheet, a, b)
    }
}

fn format_a1(row: u32, col: u32) -> String {
    let mut n = col + 1;
    let mut letters = String::new();
    while n > 0 {
        let digit = ((n - 1) % 26) as u8;
        letters.push((b'A' + digit) as char);
        n = (n - 1) / 26;
    }
    format!("{}{}", letters.chars().rev().collect::<String>(), row + 1)
}

pub fn parse_reference(source: &str, current: &CellAddress, a1: bool) -> KernelResult<RangeRef> {
    current.validate()?;
    if a1 {
        return match parse(source, current)? {
            Expr::Reference(range) => Ok(range),
            _ => Err(err("Expected a reference expression")),
        };
    }
    let text = source
        .trim()
        .strip_prefix('=')
        .unwrap_or(source.trim())
        .trim();
    let (sheet, body) = split_sheet(text)?;
    parse_r1c1(
        body,
        sheet.unwrap_or_else(|| current.sheet_id.clone()),
        current,
    )
}
fn split_sheet(text: &str) -> KernelResult<(Option<String>, &str)> {
    if let Some(tail) = text.strip_prefix('\'') {
        let mut sheet = String::new();
        let mut iter = tail.char_indices().peekable();
        while let Some((index, c)) = iter.next() {
            if c == '\'' {
                if iter.peek().is_some_and(|(_, c)| *c == '\'') {
                    iter.next();
                    sheet.push('\'');
                    continue;
                }
                let rest = &tail[index + 1..];
                if sheet.is_empty() || !rest.starts_with('!') {
                    return Err(err("Invalid quoted worksheet qualifier"));
                }
                return Ok((Some(sheet), &rest[1..]));
            }
            sheet.push(c);
        }
        return Err(err("Unterminated worksheet name"));
    }
    if let Some((sheet, body)) = text.split_once('!') {
        if sheet.is_empty()
            || !sheet
                .chars()
                .all(|c| c.is_alphanumeric() || matches!(c, '_' | '.'))
        {
            return Err(err("Invalid worksheet qualifier"));
        }
        Ok((Some(sheet.into()), body))
    } else {
        Ok((None, text))
    }
}

fn parse_r1c1(body: &str, sheet: String, current: &CellAddress) -> KernelResult<RangeRef> {
    let parts: Vec<&str> = body.split(':').collect();
    if parts.len() > 2 {
        return Err(err("Invalid R1C1 range"));
    }
    let a = endpoint_r1c1(parts[0], current)?;
    let b = endpoint_r1c1(parts.get(1).copied().unwrap_or(parts[0]), current)?;
    let r = RangeRef {
        sheet_id: sheet,
        start_row: a.0.min(b.0),
        end_row: a.0.max(b.0),
        start_column: a.1.min(b.1),
        end_column: a.1.max(b.1),
    };
    r.validate()?;
    Ok(r)
}
fn endpoint_r1c1(s: &str, current: &CellAddress) -> KernelResult<(u32, u32)> {
    let x = s.trim();
    let upper = x.to_ascii_uppercase();
    let rpos = upper
        .find('R')
        .ok_or_else(|| err("Invalid R1C1 reference"))?;
    let cpos = upper
        .find('C')
        .ok_or_else(|| err("Invalid R1C1 reference"))?;
    if rpos != 0 || cpos == 0 {
        return Err(err("Invalid R1C1 reference"));
    }
    let row = r1c1_part(&upper[1..cpos], current.row, "row")?;
    let col = r1c1_part(&upper[cpos + 1..], current.column, "column")?;
    Ok((row, col))
}
fn r1c1_part(s: &str, base: u32, kind: &str) -> KernelResult<u32> {
    if s.is_empty() {
        return Ok(base);
    }
    if s.starts_with('[') && s.ends_with(']') {
        let n: i64 = s[1..s.len() - 1]
            .parse()
            .map_err(|_| err(format!("Invalid R1C1 {kind}")))?;
        let v = (base as i64)
            .checked_add(n)
            .ok_or_else(|| err(format!("R1C1 {kind} is outside worksheet bounds")))?;
        let max = if kind == "row" {
            MAX_ROWS as i64
        } else {
            MAX_COLUMNS as i64
        };
        if v < 0 || v >= max {
            return Err(err(format!("R1C1 {kind} is outside worksheet bounds")));
        }
        return Ok(v as u32);
    }
    let n: u64 = s.parse().map_err(|_| err(format!("Invalid R1C1 {kind}")))?;
    if n == 0 {
        return Err(err(format!("Invalid R1C1 {kind}")));
    }
    let v = n - 1;
    let max = if kind == "row" { MAX_ROWS } else { MAX_COLUMNS };
    if v >= max as u64 {
        Err(err(format!("R1C1 {kind} is outside worksheet bounds")))
    } else {
        Ok(v as u32)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn at() -> CellAddress {
        CellAddress {
            sheet_id: "Sheet1".into(),
            row: 4,
            column: 4,
        }
    }
    #[test]
    fn parses_literals_and_formula() {
        let e = parse("=1+2*3", &at()).unwrap();
        assert!(matches!(e, Expr::Binary(ref op, _, _) if op == "+"));
    }
    #[test]
    fn parses_a1_and_arrays() {
        let e = parse("{1,2;3,4}", &at()).unwrap();
        assert!(matches!(e, Expr::Array(ref r) if r.len() == 2));
        assert_eq!(
            parse_reference("'O''Brien'!$A$1:B2", &at(), true)
                .unwrap()
                .sheet_id,
            "O'Brien"
        );
    }
    #[test]
    fn rejects_invalid_inputs() {
        assert_eq!(parse("=A0", &at()).unwrap_err().code, "FORMULA_PARSE");
        assert_eq!(
            parse_reference("A1048577", &at(), true).unwrap_err().code,
            "FORMULA_PARSE"
        );
    }
    #[test]
    fn parses_r1c1_relative() {
        let r = parse_reference("R[-1]C[2]", &at(), false).unwrap();
        assert_eq!((r.start_row, r.start_column), (3, 6));
    }
    #[test]
    fn parses_unicode_text_and_sheet_names_without_panicking() {
        assert_eq!(
            parse("=\"中文\"", &at()).unwrap(),
            Expr::Scalar(Scalar::Text("中文".into()))
        );
        assert_eq!(
            parse_reference("'数据表'!A1", &at(), true)
                .unwrap()
                .sheet_id,
            "数据表"
        );
    }
}
