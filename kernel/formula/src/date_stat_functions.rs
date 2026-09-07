//! Date, statistical and non-reference advanced formula functions.
//! Reference-dependent aggregation (SUMIFS/COUNTIFS/SUBTOTAL/AGGREGATE, etc.)
//! is intentionally owned by the runtime's range evaluator.

use crate::{FormulaValue, num};
use kernel_core::{KernelResult, Scalar};

pub const FUNCTIONS: &[&str] = &[
    "DATE",
    "DATEVALUE",
    "DAY",
    "MONTH",
    "YEAR",
    "HOUR",
    "MINUTE",
    "SECOND",
    "WEEKDAY",
    "EDATE",
    "EOMONTH",
    "DAYS",
    "MEDIAN",
    "LARGE",
    "SMALL",
];

fn err(code: &str, message: impl Into<String>) -> FormulaValue {
    FormulaValue::error(code, message)
}
fn value_error(message: impl Into<String>) -> FormulaValue {
    err("#VALUE!", message)
}
fn formula_error(e: kernel_core::KernelError) -> FormulaValue {
    if e.code.starts_with('#') {
        err(&e.code, e.message)
    } else {
        value_error(e.message)
    }
}
fn one(args: &[FormulaValue], index: usize) -> Result<Scalar, FormulaValue> {
    args.get(index)
        .map(|v| v.scalar())
        .ok_or_else(|| value_error("Missing argument"))
}
fn number(args: &[FormulaValue], index: usize) -> Result<f64, FormulaValue> {
    let s = one(args, index)?;
    num(&s).map_err(formula_error)
}
fn flat(v: &FormulaValue) -> Vec<Scalar> {
    match v {
        FormulaValue::Scalar(s) => vec![s.clone()],
        FormulaValue::Array(a) => a.iter().flat_map(|r| r.iter().cloned()).collect(),
    }
}
fn numbers(v: &FormulaValue) -> Result<Vec<f64>, FormulaValue> {
    flat(v)
        .iter()
        .map(|s| num(s).map_err(formula_error))
        .collect()
}
fn all_numbers(args: &[FormulaValue]) -> Result<Vec<f64>, FormulaValue> {
    args.iter().try_fold(Vec::new(), |mut out, v| {
        out.extend(numbers(v)?);
        Ok(out)
    })
}
fn result_number(n: f64) -> FormulaValue {
    if n.is_finite() {
        FormulaValue::Scalar(Scalar::Number(n))
    } else {
        value_error("Non-finite formula result")
    }
}

#[derive(Clone, Copy)]
struct Date {
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    serial: f64,
}

// Days from civil date, relative to 1970-01-01 (Howard Hinnant's proleptic Gregorian algorithm).
fn days_from_civil(y: i32, m: u32, d: u32) -> i64 {
    let y = y as i64 - if m <= 2 { 1 } else { 0 };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400;
    let mp = m as i64 + if m > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + d as i64 - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}
fn civil_from_days(z: i64) -> (i32, u32, u32) {
    let z = z + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = mp + if mp < 10 { 3 } else { -9 };
    (y as i32 + if m <= 2 { 1 } else { 0 }, m as u32, d as u32)
}
fn epoch(system1904: bool) -> i64 {
    days_from_civil(
        if system1904 { 1904 } else { 1899 },
        if system1904 { 1 } else { 12 },
        if system1904 { 1 } else { 31 },
    )
}
fn serial_from_parts(y: i32, m: u32, d: u32, h: u32, min: u32, s: u32, system1904: bool) -> f64 {
    let days = days_from_civil(y, m, d) - epoch(system1904);
    let mut serial = days as f64 + (h * 3600 + min * 60 + s) as f64 / 86400.0;
    if !system1904 && (y > 1900 || (y == 1900 && (m > 2 || (m == 2 && d >= 29)))) {
        serial += 1.0;
    }
    serial
}
fn date_from_serial(serial: f64, system1904: bool) -> Option<Date> {
    if !serial.is_finite() {
        return None;
    }
    let whole = serial.floor();
    let fraction = serial - whole;
    if !system1904 && whole == 60.0 {
        return Some(Date {
            year: 1900,
            month: 2,
            day: 29,
            hour: (fraction * 86400.0).round() as u32 / 3600,
            minute: (fraction * 86400.0).round() as u32 / 60 % 60,
            second: (fraction * 86400.0).round() as u32 % 60,
            serial,
        });
    }
    let adjusted = if !system1904 && whole > 60.0 {
        whole - 1.0
    } else {
        whole
    };
    let (y, m, d) = civil_from_days(epoch(system1904) + adjusted as i64);
    let total = (fraction * 86400.0).round() as u32;
    Some(Date {
        year: y,
        month: m,
        day: d,
        hour: total / 3600,
        minute: total / 60 % 60,
        second: total % 60,
        serial,
    })
}
fn parse_date(v: &Scalar, system1904: bool) -> Option<Date> {
    match v {
        Scalar::Number(n) => date_from_serial(*n, system1904),
        Scalar::Text(s) => {
            let x = s.trim().split(['T', ' ']).next()?.replace('/', "-");
            let p: Vec<_> = x.split('-').collect();
            if p.len() != 3 {
                return None;
            }
            let (y, m, d) = (p[0].parse().ok()?, p[1].parse().ok()?, p[2].parse().ok()?);
            Some(Date {
                year: y,
                month: m,
                day: d,
                hour: 0,
                minute: 0,
                second: 0,
                serial: serial_from_parts(y, m, d, 0, 0, 0, system1904),
            })
        }
        _ => num(v).ok().and_then(|n| date_from_serial(n, system1904)),
    }
}
fn date_parts(args: &[FormulaValue], system1904: bool) -> Result<Date, FormulaValue> {
    let y = number(args, 0)?.trunc() as i32;
    let month = number(args, 1)?.trunc() as i64;
    let day = number(args, 2)?.trunc() as i64;
    if !(1..=9999).contains(&y) || month < -120000 || month > 120000 {
        return Err(value_error("Invalid date parameters"));
    }
    let absolute = y as i64 * 12 + month - 1;
    let yy = absolute.div_euclid(12) as i32;
    let mm = absolute.rem_euclid(12) as u32 + 1;
    let base = days_from_civil(yy, mm, 1) + day - 1;
    let (ry, rm, rd) = civil_from_days(base);
    Some(Date {
        year: ry,
        month: rm,
        day: rd,
        hour: 0,
        minute: 0,
        second: 0,
        serial: serial_from_parts(ry, rm, rd, 0, 0, 0, system1904),
    })
    .ok_or_else(|| value_error("Invalid date parameters"))
}

#[cfg(test)]
pub fn call(name: &str, args: &[FormulaValue]) -> Option<KernelResult<FormulaValue>> {
    call_with_date_system(name, args, false)
}
pub fn call_with_date_system(
    name: &str,
    args: &[FormulaValue],
    date1904: bool,
) -> Option<KernelResult<FormulaValue>> {
    let n = name.to_ascii_uppercase();
    let out = match n.as_str() {
        "DATE" => date_parts(args, date1904).map(|d| result_number(d.serial.round())),
        "DATEVALUE" => one(args, 0).and_then(|s| {
            parse_date(&s, date1904)
                .map(|d| result_number(d.serial.floor()))
                .ok_or_else(|| value_error("Cannot parse date in DATEVALUE"))
        }),
        "DAY" | "MONTH" | "YEAR" | "HOUR" | "MINUTE" | "SECOND" => one(args, 0).and_then(|s| {
            parse_date(&s, date1904)
                .map(|d| {
                    result_number(match n.as_str() {
                        "DAY" => d.day as f64,
                        "MONTH" => d.month as f64,
                        "YEAR" => d.year as f64,
                        "HOUR" => d.hour as f64,
                        "MINUTE" => d.minute as f64,
                        _ => d.second as f64,
                    })
                })
                .ok_or_else(|| value_error("Invalid date"))
        }),
        "WEEKDAY" => one(args, 0).and_then(|s| {
            let typ = if let Some(value) = args.get(1) {
                num(&value.scalar()).map_err(|e| err(&e.code, e.message))? as i32
            } else {
                1
            };
            parse_date(&s, date1904)
                .map(|d| {
                    let days = days_from_civil(d.year, d.month, d.day);
                    let sunday = ((days + 4).rem_euclid(7)) as i32;

                    if ![1, 2, 3, 11, 12, 13, 14, 15, 16, 17].contains(&typ) {
                        return err("#NUM!", "Invalid WEEKDAY return_type");
                    }
                    result_number(if typ == 1 {
                        (sunday + 1) as f64
                    } else {
                        let mon = (sunday + 6) % 7;
                        if typ == 3 {
                            mon as f64
                        } else if typ == 2 || typ == 11 {
                            (mon + 1) as f64
                        } else {
                            ((mon - (typ - 11) + 7) % 7 + 1) as f64
                        }
                    })
                })
                .ok_or_else(|| value_error("Invalid date"))
        }),
        "EDATE" | "EOMONTH" => one(args, 0)
            .and_then(|s| parse_date(&s, date1904).ok_or_else(|| value_error("Invalid date")))
            .and_then(|d| number(args, 1).map(|months| (d, months.trunc() as i64)))
            .and_then(|(d, months)| {
                let abs = d.year as i64 * 12 + d.month as i64 - 1 + months;
                let y = abs.div_euclid(12) as i32;
                let m = abs.rem_euclid(12) as u32 + 1;
                let next = days_from_civil(y, m, 1) + 31;
                let (_, _, last) = civil_from_days(next - (civil_from_days(next).2 as i64));
                let day = if n == "EOMONTH" {
                    last
                } else {
                    d.day.min(last)
                };
                Ok(result_number(
                    serial_from_parts(y, m, day, 0, 0, 0, date1904).round(),
                ))
            }),
        "DAYS" => one(args, 0).and_then(|a| {
            one(args, 1).and_then(|b| {
                let x = parse_date(&a, date1904);
                let y = parse_date(&b, date1904);
                match (x, y) {
                    (Some(x), Some(y)) => Ok(result_number((x.serial - y.serial).round())),
                    _ => Err(value_error("Invalid dates in DAYS")),
                }
            })
        }),
        "MEDIAN" => all_numbers(args).map(|mut v| {
            if v.is_empty() {
                err("#NUM!", "No values for MEDIAN")
            } else {
                v.sort_by(|a, b| a.partial_cmp(b).unwrap());
                let i = v.len() / 2;
                result_number(if v.len() % 2 == 0 {
                    (v[i - 1] + v[i]) / 2.0
                } else {
                    v[i]
                })
            }
        }),
        "LARGE" | "SMALL" => args
            .get(0)
            .ok_or_else(|| value_error("Missing array"))
            .and_then(numbers)
            .and_then(|v| number(args, 1).map(|k| (v, k)))
            .map(|(mut v, k)| {
                let k = k.trunc() as usize;
                if k == 0 || k > v.len() {
                    return err("#NUM!", "Invalid k");
                }
                v.sort_by(|a, b| a.partial_cmp(b).unwrap());
                result_number(if n == "LARGE" {
                    v[v.len() - k]
                } else {
                    v[k - 1]
                })
            }),
        _ => return None,
    };
    // Formula argument failures are formula values by contract.
    Some(Ok(out.unwrap_or_else(|e| e)))
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn dates_and_serials() {
        assert_eq!(
            call(
                "DATE",
                &[
                    FormulaValue::Scalar(Scalar::Number(1900.0)),
                    FormulaValue::Scalar(Scalar::Number(1.0)),
                    FormulaValue::Scalar(Scalar::Number(1.0))
                ]
            )
            .unwrap()
            .unwrap()
            .scalar(),
            Scalar::Number(1.0)
        );
        assert_eq!(
            call("DAY", &[FormulaValue::Scalar(Scalar::Number(60.0))])
                .unwrap()
                .unwrap()
                .scalar(),
            Scalar::Number(29.0)
        );
    }
    #[test]
    fn statistic_success_and_rejection() {
        let a = FormulaValue::Array(vec![vec![
            Scalar::Number(1.0),
            Scalar::Number(3.0),
            Scalar::Number(5.0),
        ]]);
        assert_eq!(
            call("MEDIAN", &[a]).unwrap().unwrap().scalar(),
            Scalar::Number(3.0)
        );
        let bad = call(
            "LARGE",
            &[
                FormulaValue::Array(vec![vec![Scalar::Number(1.0)]]),
                FormulaValue::Scalar(Scalar::Number(2.0)),
            ],
        )
        .unwrap()
        .unwrap()
        .scalar();
        assert!(matches!(bad, Scalar::Error(_)));
    }
}
