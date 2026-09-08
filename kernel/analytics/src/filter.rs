use crate::source::SourceIndex;
use crate::task::TaskContext;
use kernel_core::{KernelError, KernelResult, RangeRef, Scalar};
use serde::{Deserialize, Serialize};
use std::cmp::Ordering;
use std::collections::BTreeMap;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FilterOwner {
    pub id: String,
    pub range: RangeRef,
    pub columns: Vec<FilterColumn>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FilterColumn {
    pub column: u32,
    pub predicate: FilterPredicate,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum FilterPredicate {
    All,
    Not {
        predicate: Box<FilterPredicate>,
    },
    Values {
        values: Vec<Scalar>,
        include_blank: bool,
        #[serde(default)]
        date_groups: Vec<DateGroupItem>,
    },
    Custom {
        join: FilterJoin,
        conditions: Vec<FilterComparison>,
    },
    Dynamic {
        #[serde(rename = "type")]
        filter_type: DynamicFilterType,
        #[serde(default)]
        value: Option<f64>,
        #[serde(default)]
        max_value: Option<f64>,
    },
    Top10 {
        top: bool,
        percent: bool,
        rank: usize,
        #[serde(default)]
        filter_value: Option<f64>,
    },
    Color {
        target: ColorTarget,
        dxf_id: i64,
        #[serde(default)]
        style: Option<serde_json::Value>,
    },
    Icon {
        icon_set: String,
        icon_id: u32,
    },
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DateGroupItem {
    pub year: i32,
    #[serde(default)]
    pub month: Option<u8>,
    #[serde(default)]
    pub day: Option<u8>,
    #[serde(default)]
    pub hour: Option<u8>,
    #[serde(default)]
    pub minute: Option<u8>,
    #[serde(default)]
    pub second: Option<u8>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterComparison {
    pub operator: FilterOperator,
    pub value: Scalar,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum FilterJoin {
    And,
    Or,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum FilterOperator {
    Equals,
    NotEquals,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual,
    Contains,
    NotContains,
    BeginsWith,
    EndsWith,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum DynamicFilterType {
    Today,
    Yesterday,
    Tomorrow,
    ThisWeek,
    LastWeek,
    NextWeek,
    ThisMonth,
    LastMonth,
    NextMonth,
    ThisQuarter,
    LastQuarter,
    NextQuarter,
    ThisYear,
    LastYear,
    NextYear,
    YearToDate,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ColorTarget {
    Cell,
    Font,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FilterRequest {
    pub revision: u64,
    pub range: RangeRef,
    #[serde(default)]
    pub owners: Vec<FilterOwner>,
    #[serde(default)]
    pub conditions: Vec<FilterColumn>,
    #[serde(default)]
    pub sort: Vec<SortKey>,
    #[serde(default)]
    pub offset: usize,
    #[serde(default = "default_page")]
    pub limit: usize,
    #[serde(default)]
    pub budget: Option<kernel_core::TaskBudget>,
    /// Workbook reference date as an Excel serial. Dynamic date predicates
    /// are rejected when this is absent instead of consulting host time.
    #[serde(default)]
    pub now_serial: Option<f64>,
}
fn default_page() -> usize {
    1024
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SortKey {
    pub column: u32,
    #[serde(default)]
    pub descending: bool,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct VisibilityBitmap {
    pub rows: u32,
    pub hidden: Vec<u64>,
    pub reasons: BTreeMap<String, Vec<u64>>,
}
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq, PartialOrd, Ord)]
#[serde(rename_all = "camelCase")]
pub enum VisibilityReason {
    Filter,
    ManualHidden,
    OwnerConflict,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FilterResult {
    pub revision: u64,
    pub owner_id: Option<String>,
    pub visible_rows: Vec<u32>,
    pub total_rows: u32,
    pub visibility: VisibilityBitmap,
    pub domain: BTreeMap<u32, Vec<DomainValue>>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DomainValue {
    pub value: Scalar,
    pub count: u32,
}

/// Evaluate a filter from dense column vectors. Missing cells are Null and are
/// never coerced to another type while criteria are evaluated.
pub(crate) fn execute_filter(
    ctx: &TaskContext<'_>,
    r: FilterRequest,
) -> KernelResult<FilterResult> {
    r.range.validate()?;
    let owner = resolve_owner(&r.range, &r.owners)?;
    validate_columns(&r.range, owner, &r.conditions)?;
    let row_count = range_row_count(&r.range)?;
    let source = SourceIndex::build(ctx, &r.range, &[])?;
    let mut criteria = owner.map(|o| o.columns.clone()).unwrap_or_default();
    criteria.extend(r.conditions.clone());
    validate_predicates(&criteria)?;
    if criteria
        .iter()
        .any(|criterion| matches!(&criterion.predicate, FilterPredicate::Dynamic { .. }))
        && r.now_serial.is_none()
    {
        return Err(KernelError::new(
            "FILTER_DATE_REFERENCE_REQUIRED",
            "Dynamic date filters require explicit nowSerial",
        )
        .recover("provide-nowSerial"));
    }
    let masks = filter_masks(ctx, &source, &criteria, r.now_serial)?;
    let mut passed = vec![false; row_count];
    let mut hidden = vec![0u64; (row_count + 63) / 64];
    let mut reasons = BTreeMap::<String, Vec<u64>>::new();
    for index in 0..row_count {
        ctx.checkpoint("filter-evaluate")?;
        passed[index] = masks.iter().all(|mask| mask[index]);
        if !passed[index] {
            mark(&mut hidden, index);
            mark_reason(&mut reasons, VisibilityReason::Filter, index);
        }
    }
    let mut domain = BTreeMap::new();
    for column in r.range.start_column..=r.range.end_column {
        let mut counts = BTreeMap::<String, (Scalar, u32)>::new();
        for index in 0..row_count {
            if criteria
                .iter()
                .enumerate()
                .all(|(criterion_index, criterion)| {
                    criterion.column == column || masks[criterion_index][index]
                })
            {
                let value = source.value_at(column, index)?;
                let entry = counts.entry(scalar_key(&value)).or_insert((value, 0));
                entry.1 += 1;
            }
        }
        domain.insert(
            column,
            counts
                .into_values()
                .map(|(value, count)| DomainValue { value, count })
                .collect(),
        );
    }
    let mut rows: Vec<usize> = (0..row_count).filter(|index| passed[*index]).collect();
    stable_sort(&mut rows, &source, &r.sort, &r.range, ctx)?;
    let visible_rows = rows
        .into_iter()
        .skip(r.offset)
        .take(r.limit)
        .map(|index| r.range.start_row + index as u32)
        .collect();
    Ok(FilterResult {
        revision: r.revision,
        owner_id: owner.map(|o| o.id.clone()),
        visible_rows,
        total_rows: row_count as u32,
        visibility: VisibilityBitmap {
            rows: row_count as u32,
            hidden,
            reasons,
        },
        domain,
    })
}

fn range_row_count(range: &RangeRef) -> KernelResult<usize> {
    usize::try_from(u64::from(range.end_row) - u64::from(range.start_row) + 1).map_err(|_| {
        KernelError::new(
            "FILTER_RANGE_TOO_LARGE",
            "Filter range cannot be represented in memory",
        )
    })
}
fn validate_columns(
    range: &RangeRef,
    owner: Option<&FilterOwner>,
    conditions: &[FilterColumn],
) -> KernelResult<()> {
    let check = |column: u32| {
        if column < range.start_column || column > range.end_column {
            Err(KernelError::new(
                "FILTER_COLUMN_OUTSIDE_RANGE",
                "Filter column is outside the requested range",
            )
            .at(column.to_string()))
        } else {
            Ok(())
        }
    };
    if let Some(owner) = owner {
        for criterion in &owner.columns {
            check(criterion.column)?;
        }
    }
    for criterion in conditions {
        check(criterion.column)?;
    }
    Ok(())
}
fn validate_predicates(criteria: &[FilterColumn]) -> KernelResult<()> {
    for criterion in criteria {
        if let FilterPredicate::Not { predicate } = &criterion.predicate {
            if !matches!(
                predicate.as_ref(),
                FilterPredicate::All
                    | FilterPredicate::Values { .. }
                    | FilterPredicate::Custom { .. }
                    | FilterPredicate::Not { .. }
            ) {
                return Err(KernelError::new(
                    "FILTER_PREDICATE_INVALID",
                    "Negation requires a scalar value or custom predicate",
                ));
            }
            validate_predicates(&[FilterColumn {
                column: criterion.column,
                predicate: *predicate.clone(),
            }])?;
        }
        match &criterion.predicate {
            FilterPredicate::Values { date_groups, .. }
                if date_groups.iter().any(|group| !valid_date_group(group)) =>
            {
                return Err(KernelError::new(
                    "FILTER_PREDICATE_INVALID",
                    "Date group contains an invalid component",
                ));
            }
            FilterPredicate::Custom { conditions, .. }
                if conditions.is_empty() || conditions.len() > 2 =>
            {
                return Err(KernelError::new(
                    "FILTER_PREDICATE_INVALID",
                    "Custom filter requires one or two conditions",
                ));
            }
            FilterPredicate::Top10 { rank, .. } if *rank == 0 => {
                return Err(KernelError::new(
                    "FILTER_PREDICATE_INVALID",
                    "Top10 requires a positive rank",
                ));
            }
            FilterPredicate::Dynamic {
                value, max_value, ..
            } if value.is_some_and(|n| !n.is_finite())
                || max_value.is_some_and(|n| !n.is_finite()) =>
            {
                return Err(KernelError::new(
                    "FILTER_PREDICATE_INVALID",
                    "Dynamic filter bounds must be finite",
                ));
            }
            _ => {}
        }
    }
    Ok(())
}
fn resolve_owner<'a>(
    range: &RangeRef,
    owners: &'a [FilterOwner],
) -> KernelResult<Option<&'a FilterOwner>> {
    for owner in owners {
        owner.range.validate()?;
        for criterion in &owner.columns {
            if criterion.column < owner.range.start_column
                || criterion.column > owner.range.end_column
            {
                return Err(KernelError::new(
                    "FILTER_COLUMN_OUTSIDE_OWNER",
                    "Filter column is outside its owner range",
                )
                .at(&owner.id));
            }
        }
    }
    let matches: Vec<_> = owners
        .iter()
        .filter(|owner| owner.range.intersects(range))
        .collect();
    if matches.len() > 1 {
        return Err(KernelError::new(
            "FILTER_OWNER_CONFLICT",
            "Overlapping filter ranges have multiple owners",
        )
        .recover("remove-overlap"));
    }
    Ok(matches.first().copied())
}

/// Masks are indexed by filter instance, so independent criteria on the same
/// field never overwrite one another. They are also reused by pivot/query.
pub(crate) fn filter_masks(
    ctx: &TaskContext<'_>,
    source: &SourceIndex,
    criteria: &[FilterColumn],
    now_serial: Option<f64>,
) -> KernelResult<Vec<Vec<bool>>> {
    validate_predicates(criteria)?;
    let mut masks = vec![vec![true; source.row_count]; criteria.len()];
    ctx.memory(
        source.bytes + source.row_count as u64 * criteria.len() as u64,
        "filter-masks",
    )?;
    for (index, criterion) in criteria.iter().enumerate() {
        source.column(criterion.column)?;
        if matches!(criterion.predicate, FilterPredicate::Top10 { .. }) {
            continue;
        }
        for row in 0..source.row_count {
            if row % 256 == 0 {
                ctx.checkpoint("filter-mask")?;
            }
            let value = source.value_at(criterion.column, row)?;
            masks[index][row] = match &criterion.predicate {
                FilterPredicate::Dynamic { filter_type, .. } => dynamic_date_matches(
                    &value,
                    now_serial.ok_or_else(|| {
                        KernelError::new(
                            "FILTER_DATE_REFERENCE_REQUIRED",
                            "Dynamic date filters require explicit nowSerial",
                        )
                    })?,
                    filter_type,
                ),
                FilterPredicate::Color {
                    target,
                    dxf_id,
                    style,
                } => {
                    let cell = ctx.reader.read_cell(&kernel_core::CellAddress {
                        sheet_id: source.range.sheet_id.clone(),
                        row: source.range.start_row + row as u32,
                        column: criterion.column,
                    })?;
                    let metadata = cell.as_ref().map(|cell| &cell.metadata);
                    let property = match target {
                        ColorTarget::Cell => "background",
                        ColorTarget::Font => "textColor",
                    };
                    let requested = style
                        .as_ref()
                        .and_then(|style| style.get(property))
                        .and_then(serde_json::Value::as_str);
                    if let Some(requested) = requested {
                        metadata
                            .and_then(|metadata| metadata.get("style"))
                            .and_then(|style| style.get(property))
                            .and_then(serde_json::Value::as_str)
                            .is_some_and(|actual| actual.eq_ignore_ascii_case(requested))
                    } else if *dxf_id >= 0 {
                        metadata
                            .and_then(|metadata| metadata.get("filterMetadata"))
                            .and_then(|metadata| metadata.get("color"))
                            .and_then(|color| color.get("dxfId"))
                            .and_then(serde_json::Value::as_i64)
                            == Some(*dxf_id)
                    } else {
                        return Err(KernelError::new(
                            "FILTER_COLOR_INVALID",
                            "Color filter requires an explicit style color or differential style identity",
                        ));
                    }
                }
                FilterPredicate::Icon { icon_set, icon_id } => {
                    let cell = ctx.reader.read_cell(&kernel_core::CellAddress {
                        sheet_id: source.range.sheet_id.clone(),
                        row: source.range.start_row + row as u32,
                        column: criterion.column,
                    })?;
                    let icon = cell
                        .as_ref()
                        .and_then(|cell| cell.metadata.get("filterMetadata"))
                        .and_then(|metadata| metadata.get("icon"));
                    icon.is_some_and(|icon| {
                        icon.get("iconSet").and_then(serde_json::Value::as_str)
                            == Some(icon_set.as_str())
                            && icon.get("iconId").and_then(serde_json::Value::as_u64)
                                == Some(*icon_id as u64)
                    })
                }
                _ => matches_pred(&value, &criterion.predicate),
            };
        }
    }
    for (index, criterion) in criteria.iter().enumerate() {
        let FilterPredicate::Top10 {
            top, percent, rank, ..
        } = &criterion.predicate
        else {
            continue;
        };
        let mut candidates = Vec::new();
        for row in 0..source.row_count {
            if row % 256 == 0 {
                ctx.checkpoint("filter-top-domain")?;
            }
            if criteria.iter().enumerate().all(|(other, filter)| {
                filter.column == criterion.column
                    || matches!(filter.predicate, FilterPredicate::Top10 { .. })
                    || masks[other][row]
            }) {
                if let Scalar::Number(value) = source.value_at(criterion.column, row)? {
                    candidates.push((row, value));
                }
            }
        }
        candidates.sort_by(|a, b| {
            let order = a.1.total_cmp(&b.1);
            if *top { order.reverse() } else { order }
        });
        let count = if *percent {
            (candidates.len() as f64 * (*rank as f64) / 100.).ceil() as usize
        } else {
            *rank
        };
        masks[index].fill(false);
        if let Some((_, boundary)) = candidates.get(
            count
                .saturating_sub(1)
                .min(candidates.len().saturating_sub(1)),
        ) {
            for (row, value) in &candidates {
                if if *top {
                    value >= boundary
                } else {
                    value <= boundary
                } {
                    masks[index][*row] = true;
                }
            }
        }
    }
    Ok(masks)
}

pub(crate) fn matches_pred(value: &Scalar, predicate: &FilterPredicate) -> bool {
    match predicate {
        FilterPredicate::All => true,
        FilterPredicate::Not { predicate } => !matches_pred(value, predicate),
        FilterPredicate::Values {
            values,
            include_blank,
            date_groups,
        } => {
            (*include_blank && matches!(value, Scalar::Null))
                || values.iter().any(|candidate| candidate == value)
                || matches!(value, Scalar::Number(serial) if serial.is_finite() && date_groups.iter().any(|group| date_group_matches(*serial, group)))
        }
        FilterPredicate::Custom { join, conditions } => {
            let mut iter = conditions
                .iter()
                .map(|condition| comparison_matches(value, condition));
            match join {
                FilterJoin::And => iter.all(|matched| matched),
                FilterJoin::Or => iter.any(|matched| matched),
            }
        }
        FilterPredicate::Dynamic { .. }
        | FilterPredicate::Top10 { .. }
        | FilterPredicate::Color { .. }
        | FilterPredicate::Icon { .. } => false,
    }
}

fn comparison_matches(value: &Scalar, comparison: &FilterComparison) -> bool {
    match comparison.operator {
        FilterOperator::Equals => return value == &comparison.value,
        FilterOperator::NotEquals => return value != &comparison.value,
        FilterOperator::LessThan
        | FilterOperator::LessThanOrEqual
        | FilterOperator::GreaterThan
        | FilterOperator::GreaterThanOrEqual => {
            if std::mem::discriminant(value) != std::mem::discriminant(&comparison.value) {
                return false;
            }
            let order = crate::aggregate::compare_scalar(value, &comparison.value);
            return match comparison.operator {
                FilterOperator::LessThan => order.is_lt(),
                FilterOperator::LessThanOrEqual => !order.is_gt(),
                FilterOperator::GreaterThan => order.is_gt(),
                FilterOperator::GreaterThanOrEqual => !order.is_lt(),
                _ => unreachable!(),
            };
        }
        _ => {}
    }
    let text = match value {
        Scalar::Text(value) => value.as_str(),
        _ => {
            return match comparison.operator {
                FilterOperator::Equals => value == &comparison.value,
                FilterOperator::NotEquals => value != &comparison.value,
                _ => false,
            };
        }
    };
    let expected = match &comparison.value {
        Scalar::Text(value) => value.as_str(),
        _ => return false,
    };
    match comparison.operator {
        FilterOperator::Equals => value == &comparison.value,
        FilterOperator::NotEquals => value != &comparison.value,
        FilterOperator::Contains => text.contains(expected),
        FilterOperator::NotContains => !text.contains(expected),
        FilterOperator::BeginsWith => text.starts_with(expected),
        FilterOperator::EndsWith => text.ends_with(expected),
        _ => false,
    }
}

fn valid_date_group(group: &DateGroupItem) -> bool {
    group.year >= 1
        && group.month.is_none_or(|value| (1..=12).contains(&value))
        && group.day.is_none_or(|value| (1..=31).contains(&value))
        && group.hour.is_none_or(|value| value < 24)
        && group.minute.is_none_or(|value| value < 60)
        && group.second.is_none_or(|value| value < 60)
}
fn date_group_matches(serial: f64, group: &DateGroupItem) -> bool {
    let (year, month, day, hour, minute, second) = excel_serial_parts(serial);
    year == group.year
        && group.month.is_none_or(|value| value as i32 == month)
        && group.day.is_none_or(|value| value as i32 == day)
        && group.hour.is_none_or(|value| value as i32 == hour)
        && group.minute.is_none_or(|value| value as i32 == minute)
        && group.second.is_none_or(|value| value as i32 == second)
}
fn dynamic_date_matches(value: &Scalar, now: f64, filter_type: &DynamicFilterType) -> bool {
    let Scalar::Number(serial) = value else {
        return false;
    };
    if !serial.is_finite() || !now.is_finite() {
        return false;
    }
    let (year, month, day, _, _, _) = excel_serial_parts(*serial);
    let (today_year, today_month, today_day, _, _, _) = excel_serial_parts(now.floor());
    let today = now.floor();
    let day_index = today as i64;
    let week_start = today - ((day_index - 1).rem_euclid(7) as f64);
    let month_index = today_year * 12 + today_month;
    let value_month_index = year * 12 + month;
    let quarter_start = today_month - ((today_month - 1) % 3);
    let quarter_index = today_year * 4 + (quarter_start - 1) / 3;
    let value_quarter_index = year * 4 + (month - 1) / 3;
    match filter_type {
        DynamicFilterType::Today => *serial >= today && *serial < today + 1.0,
        DynamicFilterType::Yesterday => *serial >= today - 1.0 && *serial < today,
        DynamicFilterType::Tomorrow => *serial >= today + 1.0 && *serial < today + 2.0,
        DynamicFilterType::ThisWeek => *serial >= week_start && *serial < week_start + 7.0,
        DynamicFilterType::LastWeek => *serial >= week_start - 7.0 && *serial < week_start,
        DynamicFilterType::NextWeek => *serial >= week_start + 7.0 && *serial < week_start + 14.0,
        DynamicFilterType::ThisMonth => value_month_index == month_index,
        DynamicFilterType::LastMonth => value_month_index == month_index - 1,
        DynamicFilterType::NextMonth => value_month_index == month_index + 1,
        DynamicFilterType::ThisQuarter => value_quarter_index == quarter_index,
        DynamicFilterType::LastQuarter => value_quarter_index == quarter_index - 1,
        DynamicFilterType::NextQuarter => value_quarter_index == quarter_index + 1,
        DynamicFilterType::ThisYear => year == today_year,
        DynamicFilterType::LastYear => year == today_year - 1,
        DynamicFilterType::NextYear => year == today_year + 1,
        DynamicFilterType::YearToDate => {
            year == today_year
                && (month < today_month || (month == today_month && day <= today_day))
        }
    }
}
fn excel_serial_parts(serial: f64) -> (i32, i32, i32, i32, i32, i32) {
    if serial.floor() == 60. {
        let seconds = ((serial - 60.) * 86400.).round() as i32;
        return (
            1900,
            2,
            29,
            seconds / 3600,
            (seconds % 3600) / 60,
            seconds % 60,
        );
    }
    // Convert Excel serial to Unix day before civil-date decomposition.
    let days = serial.floor() as i64 - 25_569 + if serial < 60.0 { 1 } else { 0 };
    let z = days + 719468;
    let era = if z >= 0 { z } else { z - 146096 } / 146097;
    let doe = z - era * 146097;
    let yoe = (doe - doe / 1460 + doe / 36524 - doe / 146096) / 365;
    let year = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = mp + if mp < 10 { 3 } else { -9 };
    let year = year + if month <= 2 { 1 } else { 0 };
    let fraction = (serial - serial.floor()) * 86400.0;
    (
        year as i32,
        month as i32,
        day as i32,
        (fraction / 3600.0) as i32,
        ((fraction % 3600.0) / 60.0) as i32,
        (fraction % 60.0) as i32,
    )
}
fn stable_sort(
    rows: &mut [usize],
    columns: &SourceIndex,
    keys: &[SortKey],
    range: &RangeRef,
    ctx: &TaskContext<'_>,
) -> KernelResult<()> {
    for key in keys {
        if key.column < range.start_column || key.column > range.end_column {
            return Err(KernelError::new(
                "FILTER_SORT_COLUMN_OUTSIDE_RANGE",
                "Sort column is outside the requested range",
            )
            .at(key.column.to_string()));
        }
    }
    rows.sort_by(|a, b| {
        for key in keys {
            let vector = columns.column(key.column).expect("validated sort column");
            let order = scalar_cmp(&vector.value_at(*a), &vector.value_at(*b));
            if order != Ordering::Equal {
                return if key.descending {
                    order.reverse()
                } else {
                    order
                };
            }
        }
        a.cmp(b)
    });
    ctx.checkpoint("filter-sort")
}
fn scalar_key(value: &Scalar) -> String {
    serde_json::to_string(value).expect("validated Scalar serializes")
}
fn scalar_cmp(a: &Scalar, b: &Scalar) -> Ordering {
    fn kind(value: &Scalar) -> u8 {
        match value {
            Scalar::Null => 0,
            Scalar::Boolean(_) => 1,
            Scalar::Number(_) => 2,
            Scalar::Text(_) => 3,
            Scalar::Error(_) => 4,
        }
    }
    match (a, b) {
        (Scalar::Number(x), Scalar::Number(y)) => x.partial_cmp(y).unwrap_or(Ordering::Equal),
        (Scalar::Text(x), Scalar::Text(y)) => x.cmp(y),
        (Scalar::Boolean(x), Scalar::Boolean(y)) => x.cmp(y),
        (Scalar::Error(x), Scalar::Error(y)) => {
            x.code.cmp(&y.code).then_with(|| x.message.cmp(&y.message))
        }
        _ => kind(a).cmp(&kind(b)),
    }
}
fn mark(bits: &mut [u64], index: usize) {
    bits[index / 64] |= 1u64 << (index % 64);
}
fn mark_reason(map: &mut BTreeMap<String, Vec<u64>>, reason: VisibilityReason, index: usize) {
    let bits = map.entry(format!("{:?}", reason)).or_default();
    if bits.len() <= index / 64 {
        bits.resize(index / 64 + 1, 0);
    }
    mark(bits, index);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accepts_core_typed_criteria() {
        let value: FilterPredicate = serde_json::from_value(serde_json::json!({
            "kind": "values", "values": ["Open", null], "includeBlank": true,
            "dateGroups": [{"year": 2026, "month": 9}]
        }))
        .expect("core values criterion must decode");
        assert!(matches!(
            value,
            FilterPredicate::Values {
                include_blank: true,
                ..
            }
        ));
        let custom: FilterPredicate = serde_json::from_value(serde_json::json!({
            "kind": "custom", "join": "and", "conditions": [{"operator": "contains", "value": "ok"}]
        }))
        .expect("core custom criterion must decode");
        validate_predicates(&[FilterColumn {
            column: 0,
            predicate: custom,
        }])
        .expect("custom criterion is valid");
    }

    #[test]
    fn rejects_invalid_top10_and_date_group() {
        let invalid_top = FilterPredicate::Top10 {
            top: true,
            percent: false,
            rank: 0,
            filter_value: None,
        };
        assert!(
            validate_predicates(&[FilterColumn {
                column: 0,
                predicate: invalid_top
            }])
            .is_err()
        );
        let invalid_date = FilterPredicate::Values {
            values: Vec::new(),
            include_blank: false,
            date_groups: vec![DateGroupItem {
                year: 2026,
                month: Some(13),
                day: None,
                hour: None,
                minute: None,
                second: None,
            }],
        };
        assert!(
            validate_predicates(&[FilterColumn {
                column: 0,
                predicate: invalid_date
            }])
            .is_err()
        );
    }
}
