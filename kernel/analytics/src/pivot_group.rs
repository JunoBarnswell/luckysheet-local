use kernel_core::{KernelError, KernelResult, Scalar};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum DateGroupUnit {
    Year,
    Quarter,
    Month,
    Week,
    Day,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum MemberValue {
    Null,
    Boolean(bool),
    Number(f64),
    Text(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum MemberType {
    Blank,
    Boolean,
    Number,
    Text,
    Error,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct MemberKey {
    #[serde(rename = "type")]
    pub member_type: MemberType,
    pub value: MemberValue,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ManualGroup {
    pub group_id: String,
    pub name: String,
    pub items: Vec<MemberKey>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", rename_all_fields = "camelCase", deny_unknown_fields)]
pub enum Group {
    Date {
        unit: DateGroupUnit,
        #[serde(default)]
        units: Vec<DateGroupUnit>,
        start_of_week: Option<u8>,
        start: Option<Scalar>,
        end: Option<Scalar>,
        #[serde(default)]
        auto_start: bool,
        #[serde(default)]
        auto_end: bool,
    },
    Number {
        interval: f64,
        start: Option<f64>,
        end: Option<f64>,
        #[serde(default)]
        auto_start: bool,
        #[serde(default)]
        auto_end: bool,
    },
    Manual {
        groups: Vec<ManualGroup>,
    },
}

pub fn typed_member(value: &Scalar) -> MemberKey {
    match value {
        Scalar::Null => MemberKey {
            member_type: MemberType::Blank,
            value: MemberValue::Null,
        },
        Scalar::Text(v) if v.is_empty() => MemberKey {
            member_type: MemberType::Blank,
            value: MemberValue::Null,
        },
        Scalar::Boolean(v) => MemberKey {
            member_type: MemberType::Boolean,
            value: MemberValue::Boolean(*v),
        },
        Scalar::Number(v) => MemberKey {
            member_type: MemberType::Number,
            value: MemberValue::Number(*v),
        },
        Scalar::Text(v) => MemberKey {
            member_type: MemberType::Text,
            value: MemberValue::Text(v.clone()),
        },
        Scalar::Error(e) => MemberKey {
            member_type: MemberType::Error,
            value: MemberValue::Text(e.code.clone()),
        },
    }
}

pub fn group(value: &Scalar, spec: &Group) -> KernelResult<Scalar> {
    if matches!(value, Scalar::Null) || matches!(value, Scalar::Text(v) if v.is_empty()) {
        return Ok(value.clone());
    }
    match spec {
        Group::Manual { groups } => {
            let key = typed_member(value);
            Ok(groups
                .iter()
                .find(|g| g.items.iter().any(|item| item == &key))
                .map(|g| Scalar::Text(g.name.clone()))
                .unwrap_or_else(|| value.clone()))
        }
        Group::Number {
            interval,
            start,
            end,
            ..
        } => {
            if !interval.is_finite() || *interval <= 0.0 {
                return Err(KernelError::new(
                    "PIVOT_GROUP_INTERVAL_INVALID",
                    "Number grouping interval must be positive",
                ));
            }
            let Scalar::Number(number) = value else {
                return Ok(value.clone());
            };
            if !number.is_finite() {
                return Err(KernelError::new(
                    "PIVOT_GROUP_VALUE_INVALID",
                    "Number grouping requires a finite number",
                ));
            }
            let begin = start.unwrap_or(0.0);
            let result = begin + ((*number - begin) / interval).floor() * interval;
            Ok(end.map_or(Scalar::Number(result), |bound| {
                Scalar::Number(result.min(bound))
            }))
        }
        Group::Date {
            unit,
            units,
            start_of_week,
            start,
            end,
            auto_start,
            auto_end,
        } => {
            // The TypeScript engine leaves values it cannot interpret as dates
            // untouched; only the grouping specification itself is rejected.
            let Some(date) = parse_date(value) else {
                return Ok(value.clone());
            };
            let lower = start.as_ref().and_then(parse_date);
            let upper = end.as_ref().and_then(parse_date);
            if lower.as_ref().is_some_and(|d| date < *d) {
                return Ok(if *auto_start {
                    date_label(lower.as_ref().unwrap(), unit, units, *start_of_week)
                } else {
                    value.clone()
                });
            }
            if upper.as_ref().is_some_and(|d| date > *d) {
                return Ok(if *auto_end {
                    date_label(upper.as_ref().unwrap(), unit, units, *start_of_week)
                } else {
                    value.clone()
                });
            }
            Ok(date_label(&date, unit, units, *start_of_week))
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
struct Date {
    year: i32,
    month: u32,
    day: u32,
}

fn parse_date(value: &Scalar) -> Option<Date> {
    match value {
        Scalar::Number(n) if n.is_finite() => Some(from_days((*n).floor() as i64 - 25_569)),
        Scalar::Text(text) => {
            let date = text.get(..10)?;
            let mut parts = date.split('-');
            Some(Date {
                year: parts.next()?.parse().ok()?,
                month: parts.next()?.parse().ok()?,
                day: parts.next()?.parse().ok()?,
            })
        }
        _ => None,
    }
}

fn date_label(
    date: &Date,
    unit: &DateGroupUnit,
    units: &[DateGroupUnit],
    week_start: Option<u8>,
) -> Scalar {
    let selected = if units.is_empty() {
        vec![*unit]
    } else {
        units.to_vec()
    };
    let labels: Vec<String> = selected
        .iter()
        .map(|u| match u {
            DateGroupUnit::Year => date.year.to_string(),
            DateGroupUnit::Quarter => format!("{} Q{}", date.year, (date.month - 1) / 3 + 1),
            DateGroupUnit::Month => format!("{}-{:02}", date.year, date.month),
            DateGroupUnit::Week => format!("W{}", week_number(date, week_start.unwrap_or(0))),
            DateGroupUnit::Day => format!("{}-{:02}-{:02}", date.year, date.month, date.day),
        })
        .collect();
    if selected.len() == 1 && selected[0] == DateGroupUnit::Year {
        Scalar::Number(date.year as f64)
    } else {
        Scalar::Text(labels.join(" / "))
    }
}

fn week_number(date: &Date, start: u8) -> u32 {
    let jan1 = Date {
        year: date.year,
        month: 1,
        day: 1,
    };
    let offset = (weekday(&jan1) + 7 - (start % 7)) % 7;
    ((days_between(&jan1, date) + offset as i64) / 7 + 1) as u32
}

fn days_between(a: &Date, b: &Date) -> i64 {
    to_days(*b) - to_days(*a)
}
fn weekday(date: &Date) -> u8 {
    ((to_days(*date) + 4).rem_euclid(7)) as u8
}
fn to_days(d: Date) -> i64 {
    let y = d.year as i64 - (d.month <= 2) as i64;
    let era = (if y >= 0 { y } else { y - 399 }) / 400;
    let yoe = y - era * 400;
    let mp = d.month as i64 + if d.month > 2 { -3 } else { 9 };
    let doy = (153 * mp + 2) / 5 + d.day as i64 - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146097 + doe - 719468
}
fn from_days(z: i64) -> Date {
    let z = z + 719468;
    let era = (if z >= 0 { z } else { z - 146096 }) / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    Date {
        year: (y + (mp + 2) / 12) as i32,
        month: (mp + if mp < 10 { 3 } else { -9 }) as u32,
        day: (doy - (153 * mp + 2) / 5 + 1) as u32,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn preserves_typed_members() {
        assert_ne!(
            typed_member(&Scalar::Number(1.0)),
            typed_member(&Scalar::Text("1".into()))
        );
    }
    #[test]
    fn rejects_invalid_interval_and_numeric_text() {
        let g = Group::Number {
            interval: 0.0,
            start: None,
            end: None,
            auto_start: false,
            auto_end: false,
        };
        assert_eq!(
            group(&Scalar::Number(1.0), &g).unwrap_err().code,
            "PIVOT_GROUP_INTERVAL_INVALID"
        );
        let g = Group::Number {
            interval: 1.0,
            start: None,
            end: None,
            auto_start: false,
            auto_end: false,
        };
        assert_eq!(
            group(&Scalar::Text("1".into()), &g).unwrap(),
            Scalar::Text("1".into())
        );
    }
}
