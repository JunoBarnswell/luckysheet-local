use crate::{FormulaValue, num, text};
use kernel_core::{KernelError, KernelResult, Scalar};

/// Scalar functions implemented by the canonical formula runtime.
/// Aggregate and lazy functions are owned by the evaluator and intentionally
/// do not appear here.
pub const FUNCTIONS: &[&str] = &[
    "ABS",
    "SQRT",
    "POWER",
    "MOD",
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "INT",
    "TRUNC",
    "CEILING",
    "FLOOR",
    "PI",
    "EXP",
    "LN",
    "LOG",
    "LOG10",
    "SIGN",
    "CONCAT",
    "CONCATENATE",
    "TEXTJOIN",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "LOWER",
    "UPPER",
    "PROPER",
    "TRIM",
    "CLEAN",
    "EXACT",
    "FIND",
    "SEARCH",
    "REPLACE",
    "SUBSTITUTE",
    "REPT",
    "TEXT",
    "VALUE",
    "CHAR",
    "CODE",
    "ISBLANK",
    "ISNUMBER",
    "ISTEXT",
    "ISNONTEXT",
    "ISLOGICAL",
    "ISERROR",
    "ISERR",
    "ISNA",
    "N",
    "T",
    "IFNA",
    "NOT",
    "XOR",
    "SWITCH",
    "TRUE",
    "FALSE",
];

fn ok(v: Scalar) -> KernelResult<FormulaValue> {
    Ok(FormulaValue::Scalar(v))
}
fn excel(code: &str, message: impl Into<String>) -> KernelResult<FormulaValue> {
    ok(Scalar::error(code, message))
}
fn arg<'a>(args: &'a [FormulaValue], i: usize) -> KernelResult<&'a FormulaValue> {
    args.get(i)
        .ok_or_else(|| KernelError::new("FORMULA_VALUE", "Missing argument"))
}
fn scalar(v: &FormulaValue) -> Scalar {
    v.scalar()
}
fn flat(args: &[FormulaValue]) -> Vec<Scalar> {
    args.iter()
        .flat_map(|v| match v {
            FormulaValue::Scalar(s) => vec![s.clone()],
            FormulaValue::Array(a) => a.iter().flat_map(|r| r.clone()).collect(),
        })
        .collect()
}
fn string(v: &Scalar) -> String {
    text(v)
}
fn number(v: &Scalar) -> KernelResult<f64> {
    num(v)
}

fn require(args: &[FormulaValue], n: usize) -> KernelResult<()> {
    if args.len() < n {
        Err(KernelError::new("FORMULA_VALUE", "Missing argument"))
    } else {
        Ok(())
    }
}
fn rounded(n: f64, digits: f64, direction: i32) -> f64 {
    let p = 10f64.powf(digits.trunc());
    let x = n * p;
    let y = if direction < 0 {
        x.floor()
    } else if direction > 0 {
        x.ceil()
    } else {
        x.round()
    };
    y / p
}
fn each_cell(v: &FormulaValue) -> Box<dyn Iterator<Item = &Scalar> + '_> {
    match v {
        FormulaValue::Scalar(s) => Box::new(std::iter::once(s)),
        FormulaValue::Array(a) => Box::new(a.iter().flat_map(|r| r.iter())),
    }
}

pub fn call(name: &str, args: &[FormulaValue]) -> Option<KernelResult<FormulaValue>> {
    let n = name.to_ascii_uppercase();
    if !FUNCTIONS.contains(&n.as_str()) {
        return None;
    }
    let result = match n.as_str() {
        "ABS" => unary_num(args, |x| x.abs()),
        "SQRT" => unary_num_code(
            args,
            "#NUM!",
            "Negative number in SQRT",
            |x| x.sqrt(),
            |x| x >= 0.0,
        ),
        "POWER" => binary_num(args, |a, b| a.powf(b)),
        "MOD" => require(args, 2).and_then(|_| {
            let a = number(&scalar(&args[0]))?;
            let b = number(&scalar(&args[1]))?;
            if b == 0.0 {
                excel("#DIV/0!", "Division by zero in MOD")
            } else {
                ok(Scalar::Number(((a % b) + b) % b))
            }
        }),
        "ROUND" | "ROUNDUP" | "ROUNDDOWN" => require(args, 1).and_then(|_| {
            let n = number(&scalar(&args[0]))?;
            let d = if let Some(v) = args.get(1) {
                number(&scalar(v))?
            } else {
                0.0
            };
            let dir = if n >= 0.0 && name.eq_ignore_ascii_case("ROUNDUP") {
                1
            } else if n < 0.0 && name.eq_ignore_ascii_case("ROUNDUP") {
                -1
            } else if name.eq_ignore_ascii_case("ROUNDDOWN") {
                if n >= 0.0 { -1 } else { 1 }
            } else {
                0
            };
            ok(Scalar::Number(rounded(n, d, dir)))
        }),
        "INT" => unary_num(args, f64::floor),
        "TRUNC" => require(args, 1).and_then(|_| {
            let n = number(&scalar(&args[0]))?;
            let d = if let Some(v) = args.get(1) {
                number(&scalar(v))?
            } else {
                0.0
            };
            let p = 10f64.powf(d.trunc());
            ok(Scalar::Number((n * p).trunc() / p))
        }),
        "CEILING" | "FLOOR" => require(args, 1).and_then(|_| {
            let n = number(&scalar(&args[0]))?;
            let s = if let Some(v) = args.get(1) {
                number(&scalar(v))?
            } else {
                1.0
            };
            if s == 0.0 {
                return ok(Scalar::Number(0.0));
            }
            let q = n / s;
            let z = if n.is_sign_positive() == s.is_sign_positive() || n == 0.0 {
                if n.is_sign_positive() && name.eq_ignore_ascii_case("FLOOR") {
                    q.floor()
                } else if !n.is_sign_positive() && name.eq_ignore_ascii_case("CEILING") {
                    q.ceil()
                } else if name.eq_ignore_ascii_case("CEILING") {
                    q.ceil()
                } else {
                    q.floor()
                }
            } else {
                return excel("#NUM!", "Sign mismatch in CEILING/FLOOR");
            };
            ok(Scalar::Number(z * s))
        }),
        "PI" => ok(Scalar::Number(std::f64::consts::PI)),
        "EXP" => unary_num(args, f64::exp),
        "LN" => unary_num_code(
            args,
            "#NUM!",
            "Number must be positive in LN",
            f64::ln,
            |x| x > 0.0,
        ),
        "LOG" => require(args, 1).and_then(|_| {
            let x = number(&scalar(&args[0]))?;
            let b = if let Some(v) = args.get(1) {
                number(&scalar(v))?
            } else {
                10.0
            };
            if x <= 0.0 || b <= 0.0 || b == 1.0 {
                excel("#NUM!", "Invalid base or number in LOG")
            } else {
                ok(Scalar::Number(x.log(b)))
            }
        }),
        "LOG10" => unary_num_code(
            args,
            "#NUM!",
            "Number must be positive in LOG10",
            f64::log10,
            |x| x > 0.0,
        ),
        "SIGN" => unary_num(args, |x| {
            if x > 0.0 {
                1.0
            } else if x < 0.0 {
                -1.0
            } else {
                0.0
            }
        }),
        "CONCAT" => concat(args, true),
        "CONCATENATE" => concat(args, false),
        "TEXTJOIN" => textjoin(args),
        "LEFT" | "RIGHT" | "MID" => slice_text(&n, args),
        "LEN" => require(args, 1).and_then(|_| {
            ok(Scalar::Number(
                string(&scalar(&args[0])).chars().count() as f64
            ))
        }),
        "LOWER" => map_text(args, |s| s.to_lowercase()),
        "UPPER" => map_text(args, |s| s.to_uppercase()),
        "PROPER" => map_text(args, |s| {
            s.split_whitespace()
                .map(|w| {
                    let mut c = w.chars();
                    match c.next() {
                        Some(x) if x.is_alphabetic() => {
                            x.to_uppercase().collect::<String>() + &c.as_str().to_lowercase()
                        }
                        _ => w.to_string(),
                    }
                })
                .collect::<Vec<_>>()
                .join(" ")
        }),
        "TRIM" => map_text(args, |s| s.split_whitespace().collect::<Vec<_>>().join(" ")),
        "CLEAN" => map_text(args, |s| s.chars().filter(|c| !c.is_control()).collect()),
        "EXACT" => require(args, 2).and_then(|_| {
            ok(Scalar::Boolean(
                string(&scalar(&args[0])) == string(&scalar(&args[1])),
            ))
        }),
        "FIND" | "SEARCH" => find_text(&n, args),
        "REPLACE" => replace_text(args),
        "SUBSTITUTE" => substitute(args),
        "REPT" => rept(args),
        "TEXT" => text_fn(args),
        "VALUE" => value_fn(args),
        "CHAR" => char_fn(args),
        "CODE" => code_fn(args),
        "ISBLANK" => info(args, |v| match v {
            Scalar::Null => true,
            Scalar::Text(s) => s.is_empty(),
            _ => false,
        }),
        "ISNUMBER" => info(args, |v| matches!(v,Scalar::Number(n) if n.is_finite())),
        "ISTEXT" => info(args, |v| matches!(v, Scalar::Text(_))),
        "ISNONTEXT" => info(args, |v| !matches!(v, Scalar::Text(_))),
        "ISLOGICAL" => info(args, |v| matches!(v, Scalar::Boolean(_))),
        "ISERROR" => info(args, |v| matches!(v, Scalar::Error(_))),
        "ISERR" => info(args, |v| matches!(v,Scalar::Error(e) if e.code!="#N/A")),
        "ISNA" => info(args, |v| matches!(v,Scalar::Error(e) if e.code=="#N/A")),
        "N" => require(args, 1).and_then(|_| {
            let v = scalar(&args[0]);
            ok(Scalar::Number(match v {
                Scalar::Number(n) => n,
                Scalar::Boolean(b) => {
                    if b {
                        1.0
                    } else {
                        0.0
                    }
                }
                _ => 0.0,
            }))
        }),
        "T" => require(args, 1).and_then(|_| {
            let v = scalar(&args[0]);
            ok(Scalar::Text(match v {
                Scalar::Text(s) => s,
                _ => String::new(),
            }))
        }),
        "IFNA" => ifna(args),
        "NOT" => require(args, 1).and_then(|_| {
            logical(&scalar(&args[0]))
                .map(|b| Scalar::Boolean(!b))
                .and_then(ok)
        }),
        "XOR" => xor(args),
        "SWITCH" => switch_fn(args),
        "TRUE" => ok(Scalar::Boolean(true)),
        "FALSE" => ok(Scalar::Boolean(false)),
        _ => unreachable!(),
    };
    Some(match result {
        Ok(v) => Ok(v),
        Err(e) if e.code == "FORMULA_VALUE" => Ok(FormulaValue::error("#VALUE!", e.message)),
        Err(e) => Err(e),
    })
}

fn unary_num<F: Fn(f64) -> f64>(a: &[FormulaValue], f: F) -> KernelResult<FormulaValue> {
    require(a, 1)
        .and_then(|_| {
            number(&scalar(&a[0]))
                .map(|x| Scalar::Number(f(x)))
                .map_err(|e| e)
        })
        .and_then(ok)
}
fn unary_num_code<F: Fn(f64) -> f64, P: Fn(f64) -> bool>(
    a: &[FormulaValue],
    c: &str,
    m: &str,
    f: F,
    p: P,
) -> KernelResult<FormulaValue> {
    require(a, 1).and_then(|_| {
        let x = number(&scalar(&a[0]))?;
        if p(x) {
            ok(Scalar::Number(f(x)))
        } else {
            excel(c, m)
        }
    })
}
fn binary_num<F: Fn(f64, f64) -> f64>(a: &[FormulaValue], f: F) -> KernelResult<FormulaValue> {
    require(a, 2)
        .and_then(|_| {
            Ok(Scalar::Number(f(
                number(&scalar(&a[0]))?,
                number(&scalar(&a[1]))?,
            )))
        })
        .and_then(ok)
}
fn concat(a: &[FormulaValue], flatten: bool) -> KernelResult<FormulaValue> {
    let s = if flatten {
        flat(a).iter().map(string).collect()
    } else {
        a.iter().map(|v| string(&scalar(v))).collect()
    };
    ok(Scalar::Text(s))
}
fn map_text<F: Fn(String) -> String>(a: &[FormulaValue], f: F) -> KernelResult<FormulaValue> {
    require(a, 1).and_then(|_| ok(Scalar::Text(f(string(&scalar(&a[0]))))))
}
fn textjoin(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 2).and_then(|_| {
        let d = string(&scalar(&a[0]));
        let ignore = logical(&scalar(&a[1]))?;
        let mut x = Vec::new();
        for v in &a[2..] {
            for s in each_cell(v) {
                let q = string(s);
                if !ignore || !q.is_empty() {
                    x.push(q)
                }
            }
        }
        ok(Scalar::Text(x.join(&d)))
    })
}
fn slice_text(n: &str, a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 1).and_then(|_| {
        let s = string(&scalar(&a[0]));
        let count = if n == "MID" {
            number(&scalar(arg(a, 2).map_err(|e| e)?))?
        } else if let Some(v) = a.get(1) {
            number(&scalar(v))?
        } else {
            1.0
        };
        if count < 0.0 {
            return excel("#VALUE!", "Invalid length");
        };
        let k = count.floor() as usize;
        if n == "MID" {
            let start = number(&scalar(&a[1]))?;
            if start < 1.0 {
                return excel("#VALUE!", "Invalid start");
            };
            ok(Scalar::Text(
                s.chars().skip(start.floor() as usize - 1).take(k).collect(),
            ))
        } else if n == "LEFT" {
            ok(Scalar::Text(s.chars().take(k).collect()))
        } else {
            let z: Vec<_> = s.chars().collect();
            ok(Scalar::Text(
                z.into_iter()
                    .rev()
                    .take(k)
                    .collect::<Vec<_>>()
                    .into_iter()
                    .rev()
                    .collect(),
            ))
        }
    })
}
fn find_text(n: &str, a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 2).and_then(|_| {
        let f = string(&scalar(&a[0]));
        let s = string(&scalar(&a[1]));
        let start = if let Some(v) = a.get(2) {
            number(&scalar(v))?
        } else {
            1.0
        };
        if start < 1.0 {
            return excel("#VALUE!", "Invalid startNum");
        };
        let (f, s) = if n == "SEARCH" {
            (f.to_lowercase(), s.to_lowercase())
        } else {
            (f, s)
        };
        match s.get(start.floor() as usize - 1..).and_then(|x| x.find(&f)) {
            Some(i) => ok(Scalar::Number((i + start.floor() as usize) as f64)),
            None => excel("#VALUE!", format!("Text not found in {n}")),
        }
    })
}
fn replace_text(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 4).and_then(|_| {
        let s = string(&scalar(&a[0]));
        let st = number(&scalar(&a[1]))?;
        let n = number(&scalar(&a[2]))?;
        if st < 1.0 || n < 0.0 {
            return excel("#VALUE!", "Invalid arguments in REPLACE");
        };
        let z: Vec<_> = s.chars().collect();
        let i = st.floor() as usize - 1;
        ok(Scalar::Text(
            z.iter().copied().take(i).collect::<String>()
                + &string(&scalar(&a[3]))
                + &z.iter()
                    .copied()
                    .skip(i + n.floor() as usize)
                    .collect::<String>(),
        ))
    })
}
fn substitute(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 3).and_then(|_| {
        let s = string(&scalar(&a[0]));
        let old = string(&scalar(&a[1]));
        let new = string(&scalar(&a[2]));
        if old.is_empty() {
            return ok(Scalar::Text(s));
        };
        if let Some(v) = a.get(3) {
            let wanted = number(&scalar(v))?.floor() as usize;
            if wanted == 0 {
                return excel("#VALUE!", "Invalid instance number");
            };
            let mut count = 0;
            let mut out = String::new();
            for p in s.split_inclusive(&old) {
                count += 1;
                if count == wanted {
                    out.push_str(&p[..p.len() - old.len()]);
                    out.push_str(&new)
                } else {
                    out.push_str(p)
                }
            }
            ok(Scalar::Text(out))
        } else {
            ok(Scalar::Text(s.replace(&old, &new)))
        }
    })
}
fn rept(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    const MAX_REPT_OUTPUT_BYTES: usize = 32_767;
    require(a, 2).and_then(|_| {
        let s = string(&scalar(&a[0]));
        let n = number(&scalar(&a[1]))?;
        if n < 0.0 || !n.is_finite() {
            return excel("#VALUE!", "Invalid count in REPT");
        };
        let repetitions = n.floor();
        if repetitions > usize::MAX as f64 {
            return excel("#VALUE!", "REPT count exceeds the supported output budget");
        }
        let repetitions = repetitions as usize;
        let output_bytes = s
            .len()
            .checked_mul(repetitions)
            .ok_or_else(|| KernelError::new("FORMULA_VALUE", "REPT output size overflow"))?;
        if output_bytes > MAX_REPT_OUTPUT_BYTES {
            return excel("#VALUE!", "REPT output exceeds the supported formula-string budget");
        }
        ok(Scalar::Text(s.repeat(repetitions)))
    })
}
fn format_number(value: f64, format: &str) -> String {
    let f = format.trim();
    if f.is_empty() || f.eq_ignore_ascii_case("general") {
        return value.to_string();
    }
    let lower = f.to_ascii_lowercase();
    if lower.contains('y') && lower.contains('d') {
        let serial = value.floor().max(1.0) as i64;
        // Civil-date conversion uses the Unix epoch; Excel serial 25569 is
        // 1970-01-01 (the 1900 leap-year bug is intentionally skipped).
        let days = serial - 25569;
        let z = days + 719468;
        let era = if z >= 0 { z } else { z - 146096 } / 146097;
        let doe = z - era * 146097;
        let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
        let y = yoe + era * 400;
        let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
        let mp = (5 * doy + 2) / 153;
        let day = doy - (153 * mp + 2) / 5 + 1;
        let month = mp + if mp < 10 { 3 } else { -9 };
        let year = y + if month <= 2 { 1 } else { 0 };
        if lower.contains("yyyy") {
            return format!("{year:04}-{month:02}-{day:02}");
        }
        return format!("{year:04}-{month:02}-{day:02}");
    }
    let scale = if f.matches('%').count() > 0 {
        100.0
    } else {
        1.0
    };
    let value = value * scale;
    let start = f.find(|c: char| matches!(c, '0' | '#' | '?')).unwrap_or(0);
    let end = f
        .rfind(|c: char| matches!(c, '0' | '#' | '?' | '.' | ','))
        .map(|i| i + 1)
        .unwrap_or(f.len());
    let pattern = &f[start..end];
    let decimals = pattern
        .split('.')
        .nth(1)
        .map(|x| x.chars().filter(|c| matches!(c, '0' | '#' | '?')).count())
        .unwrap_or(0);
    let rounded = format!("{value:.decimals$}");
    let mut parts = rounded.split('.');
    let mut integer = parts.next().unwrap_or_default().to_string();
    let fraction = parts.next().unwrap_or_default();
    if pattern.contains(',') {
        let negative = integer.starts_with('-');
        let digits = integer
            .trim_start_matches('-')
            .chars()
            .rev()
            .collect::<Vec<_>>();
        let grouped = digits
            .chunks(3)
            .map(|x| x.iter().rev().collect::<String>())
            .rev()
            .collect::<Vec<_>>()
            .join(",");
        integer = if negative {
            format!("-{grouped}")
        } else {
            grouped
        };
    }
    let number = if decimals > 0 {
        format!("{integer}.{fraction}")
    } else {
        integer
    };
    let prefix = f[..start].replace('"', "");
    let suffix = f[end..].replace('"', "");
    format!("{prefix}{number}{suffix}")
}
fn text_fn(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 2).and_then(|_| {
        let v = scalar(&a[0]);
        if let Scalar::Error(e) = v {
            return ok(Scalar::Error(e));
        };
        let format = string(&scalar(&a[1]));
        match v {
            Scalar::Number(n) => ok(Scalar::Text(format_number(n, &format))),
            _ => ok(Scalar::Text(string(&v))),
        }
    })
}
fn value_fn(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 1).and_then(|_| {
        let s = string(&scalar(&a[0]))
            .trim()
            .replace('$', "")
            .replace(',', "");
        let pct = s.ends_with('%');
        let n = s.trim_end_matches('%').parse::<f64>().map_err(|_| {
            KernelError::new("FORMULA_VALUE", "Cannot convert text to number in VALUE")
        })?;
        ok(Scalar::Number(if pct { n / 100.0 } else { n }))
    })
}
fn char_fn(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 1).and_then(|_| {
        let n = number(&scalar(&a[0]))?;
        if !(1.0..=255.0).contains(&n) {
            excel("#VALUE!", "Invalid char code")
        } else {
            ok(Scalar::Text((n as u8 as char).to_string()))
        }
    })
}
fn code_fn(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 1).and_then(|_| {
        let s = string(&scalar(&a[0]));
        s.chars()
            .next()
            .map(|c| ok(Scalar::Number(c as u32 as f64)))
            .unwrap_or_else(|| excel("#VALUE!", "Empty string in CODE"))
    })
}
fn info<F: Fn(&Scalar) -> bool>(a: &[FormulaValue], f: F) -> KernelResult<FormulaValue> {
    require(a, 1).and_then(|_| ok(Scalar::Boolean(f(&scalar(&a[0])))))
}
fn logical(v: &Scalar) -> KernelResult<bool> {
    match v {
        Scalar::Boolean(b) => Ok(*b),
        Scalar::Number(n) => Ok(*n != 0.0),
        Scalar::Text(s) if s.eq_ignore_ascii_case("TRUE") => Ok(true),
        Scalar::Text(s) if s.eq_ignore_ascii_case("FALSE") => Ok(false),
        Scalar::Error(e) => Err(KernelError::new(e.code.clone(), e.message.clone())),
        _ => Err(KernelError::new(
            "FORMULA_VALUE",
            "Expected TRUE, FALSE, number, or boolean",
        )),
    }
}
fn ifna(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    require(a, 1).and_then(|_| {
        let v = scalar(&a[0]);
        if matches!(&v,Scalar::Error(e) if e.code=="#N/A") {
            Ok(a.get(1)
                .cloned()
                .unwrap_or_else(|| FormulaValue::Scalar(Scalar::Text(String::new()))))
        } else {
            ok(v)
        }
    })
}
fn xor(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    let mut c = 0;
    for v in flat(a) {
        if logical(&v)? {
            c += 1
        }
    }
    ok(Scalar::Boolean(c % 2 == 1))
}
fn switch_fn(a: &[FormulaValue]) -> KernelResult<FormulaValue> {
    if a.len() < 3 {
        return excel("#VALUE!", "SWITCH requires target, value, result");
    };
    let target = string(&scalar(&a[0]));
    let pairs_end = if a.len() % 2 == 0 {
        a.len() - 1
    } else {
        a.len()
    };
    let mut i = 1;
    while i + 1 < pairs_end {
        if target == string(&scalar(&a[i])) {
            return Ok(a[i + 1].clone());
        }
        i += 2
    }
    if a.len() % 2 == 0 {
        Ok(a[a.len() - 1].clone())
    } else {
        excel("#N/A", "No matching case in SWITCH")
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    fn n(v: f64) -> FormulaValue {
        FormulaValue::Scalar(Scalar::Number(v))
    }
    fn s(v: &str) -> FormulaValue {
        FormulaValue::Scalar(Scalar::Text(v.into()))
    }

    #[test]
    fn text_and_math_success() {
        assert_eq!(call("POWER", &[n(2.0), n(3.0)]).unwrap().unwrap(), n(8.0));
        assert_eq!(call("CONCAT", &[s("A"), n(2.0)]).unwrap().unwrap(), s("A2"));
        assert_eq!(
            call("FIND", &[s("bc"), s("abcd")]).unwrap().unwrap(),
            n(2.0)
        );
        assert_eq!(
            call("TEXT", &[n(1234.5), s("$#,##0.00")]).unwrap().unwrap(),
            s("$1,234.50")
        );
    }

    #[test]
    fn invalid_inputs_are_excel_errors_and_unknown_is_none() {
        let value = call("SQRT", &[n(-1.0)]).unwrap().unwrap();
        assert!(matches!(value, FormulaValue::Scalar(Scalar::Error(e)) if e.code == "#NUM!"));
        let missing = call("LEFT", &[]).unwrap().unwrap();
        assert!(matches!(missing, FormulaValue::Scalar(Scalar::Error(e)) if e.code == "#VALUE!"));
        let logical_error = call("NOT", &[s("hello")]).unwrap().unwrap();
        assert!(
            matches!(logical_error, FormulaValue::Scalar(Scalar::Error(e)) if e.code == "#VALUE!")
        );
        assert!(call("DOES_NOT_EXIST", &[]).is_none());
    }

    #[test]
    fn rept_bounds_count_and_output_size_before_allocation() {
        assert_eq!(
            call("REPT", &[s("ab"), n(3.0)]).unwrap().unwrap(),
            s("ababab")
        );
        let huge = call("REPT", &[s("AA"), n(1e308)]).unwrap().unwrap();
        assert!(matches!(huge, FormulaValue::Scalar(Scalar::Error(e)) if e.code == "#VALUE!"));
        let too_large = call("REPT", &[s("x"), n(32_768.0)]).unwrap().unwrap();
        assert!(matches!(too_large, FormulaValue::Scalar(Scalar::Error(e)) if e.code == "#VALUE!"));
    }
}
