use crate::{
    aggregate::{AggregateKind, AggregateState, compare_scalar, scalar_key},
    source::SourceIndex,
    task::TaskContext,
};
use kernel_core::{KernelError, KernelResult, RangeRef, Scalar};
use serde::{Deserialize, Serialize};
use std::{
    collections::{HashMap, HashSet},
    sync::Arc,
};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct QueryRequest {
    pub revision: u64,
    pub range: RangeRef,
    #[serde(default)]
    pub filters: Vec<crate::FilterColumn>,
    #[serde(default)]
    pub joins: Vec<JoinSpec>,
    #[serde(default)]
    pub group_by: Vec<u32>,
    #[serde(default)]
    pub aggregates: Vec<QueryAggregate>,
    #[serde(default)]
    pub columns: Option<Vec<u32>>,
    #[serde(default)]
    pub sort: Vec<SortSpec>,
    pub limit: usize,
    #[serde(default)]
    pub offset: usize,
    #[serde(default)]
    pub now_serial: Option<f64>,
    #[serde(default)]
    pub budget: Option<kernel_core::TaskBudget>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct JoinSpec {
    pub range: RangeRef,
    pub left_column: u32,
    pub right_column: u32,
    pub kind: JoinKind,
    pub output_start_column: u32,
}
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum JoinKind {
    Inner,
    Left,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct QueryAggregate {
    pub column: u32,
    pub aggregate: AggregateKind,
    pub output_column: u32,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SortSpec {
    pub column: u32,
    #[serde(default)]
    pub descending: bool,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct QueryRow {
    pub row_id: u32,
    pub source_row_ids: Vec<Option<u32>>,
    pub values: Vec<Scalar>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct QueryResult {
    pub revision: u64,
    pub columns: Vec<u32>,
    pub rows: Vec<QueryRow>,
    pub total: usize,
    pub grouped: bool,
}
#[derive(Clone, Copy)]
struct ColumnBinding {
    source: usize,
    column: u32,
}
/// Input rows are represented solely by parallel source row-id vectors. Joins
/// preserve multiplicity without copying any source cell values.
struct RowPaths {
    rows: Vec<Vec<Option<usize>>>,
}
impl RowPaths {
    fn len(&self) -> usize {
        self.rows.first().map_or(0, Vec::len)
    }
}
struct GroupRow {
    keys: Vec<Scalar>,
    states: Vec<AggregateState>,
}

pub(crate) fn execute_query(ctx: &TaskContext<'_>, r: QueryRequest) -> KernelResult<QueryResult> {
    if r.limit == 0 {
        return Err(KernelError::new(
            "QUERY_PAGE_LIMIT_INVALID",
            "Query result pages require a positive limit",
        ));
    }
    let first = SourceIndex::build(ctx, &r.range, &[])?;
    let masks = crate::filter::filter_masks(ctx, &first, &r.filters, r.now_serial)?;
    let mut bindings = first
        .columns
        .iter()
        .map(|column| {
            (
                *column,
                ColumnBinding {
                    source: 0,
                    column: *column,
                },
            )
        })
        .collect::<HashMap<_, _>>();
    let mut paths = RowPaths {
        rows: vec![
            (0..first.row_count)
                .filter(|row| masks.iter().all(|mask| mask[*row]))
                .map(Some)
                .collect(),
        ],
    };
    let mut sources = vec![first];
    for join in &r.joins {
        join_rows(ctx, &mut paths, &mut sources, &mut bindings, join)?;
    }
    let source_bytes = sources.iter().map(|source| source.bytes).sum::<u64>();
    ctx.memory(
        source_bytes + paths.len() as u64 * sources.len() as u64 * 16,
        "query-rows",
    )?;
    let grouped = !r.group_by.is_empty() || !r.aggregates.is_empty();
    if grouped {
        return grouped_query(ctx, &r, &paths, &sources, &bindings, source_bytes);
    }
    let mut default_columns = bindings.keys().copied().collect::<Vec<_>>();
    default_columns.sort_unstable();
    let columns = r.columns.clone().unwrap_or(default_columns);
    for column in columns.iter().chain(r.sort.iter().map(|sort| &sort.column)) {
        binding(&bindings, *column)?;
    }
    let mut order = (0..paths.len()).collect::<Vec<_>>();
    let sort_bindings = r
        .sort
        .iter()
        .map(|sort| binding(&bindings, sort.column))
        .collect::<KernelResult<Vec<_>>>()?;
    order.sort_by(|left, right| {
        for (sort, binding) in r.sort.iter().zip(&sort_bindings) {
            let cmp = compare_scalar(
                &path_value(&paths, &sources, *left, *binding),
                &path_value(&paths, &sources, *right, *binding),
            );
            if !cmp.is_eq() {
                return if sort.descending { cmp.reverse() } else { cmp };
            }
        }
        left.cmp(right)
    });
    ctx.checkpoint("query-sorted")?;
    let total = order.len();
    let mut rows = Vec::with_capacity(r.limit.min(total.saturating_sub(r.offset)));
    for (position, row) in order.into_iter().skip(r.offset).take(r.limit).enumerate() {
        if position % 256 == 0 {
            ctx.checkpoint("query-result-page")?;
        }
        let source_row_ids = paths
            .rows
            .iter()
            .enumerate()
            .map(|(source, ids)| ids[row].map(|id| sources[source].range.start_row + id as u32))
            .collect::<Vec<_>>();
        rows.push(QueryRow {
            row_id: source_row_ids[0].expect("left source row is present"),
            source_row_ids,
            values: columns
                .iter()
                .map(|column| path_value(&paths, &sources, row, bindings[column]))
                .collect(),
        });
    }
    Ok(QueryResult {
        revision: r.revision,
        columns,
        rows,
        total,
        grouped: false,
    })
}
fn binding(bindings: &HashMap<u32, ColumnBinding>, column: u32) -> KernelResult<ColumnBinding> {
    bindings.get(&column).copied().ok_or_else(|| {
        KernelError::new(
            "QUERY_COLUMN_UNAVAILABLE",
            "Query column is not in the input schema",
        )
        .at(column.to_string())
    })
}
fn path_value(
    paths: &RowPaths,
    sources: &[Arc<SourceIndex>],
    row: usize,
    binding: ColumnBinding,
) -> Scalar {
    paths.rows[binding.source][row]
        .map(|id| {
            sources[binding.source]
                .column(binding.column)
                .expect("validated column binding")
                .value_at(id)
        })
        .unwrap_or(Scalar::Null)
}
fn join_rows(
    ctx: &TaskContext<'_>,
    paths: &mut RowPaths,
    sources: &mut Vec<Arc<SourceIndex>>,
    bindings: &mut HashMap<u32, ColumnBinding>,
    join: &JoinSpec,
) -> KernelResult<()> {
    let left = binding(bindings, join.left_column)?;
    let right = SourceIndex::build(ctx, &join.range, &[])?;
    right.column(join.right_column)?;
    let source_id = sources.len();
    let mut new_bindings = Vec::new();
    for column in &right.columns {
        let output = join
            .output_start_column
            .checked_add(column - right.range.start_column)
            .ok_or_else(|| {
                KernelError::new(
                    "QUERY_COLUMN_OVERFLOW",
                    "Join output schema column overflows",
                )
            })?;
        if bindings.contains_key(&output) {
            return Err(KernelError::new(
                "QUERY_COLUMN_COLLISION",
                "Join output columns overlap an existing input field",
            )
            .at(output.to_string()));
        }
        new_bindings.push((
            output,
            ColumnBinding {
                source: source_id,
                column: *column,
            },
        ));
    }
    let mut index = HashMap::<String, Vec<usize>>::new();
    for row in 0..right.row_count {
        if row % 256 == 0 {
            ctx.checkpoint("query-hash-join-build")?;
        }
        let value = right.value_at(join.right_column, row)?;
        if !matches!(value, Scalar::Null) {
            index.entry(scalar_key(&value)).or_default().push(row);
        }
    }
    let mut next = RowPaths {
        rows: vec![Vec::new(); sources.len() + 1],
    };
    let source_bytes = sources.iter().map(|source| source.bytes).sum::<u64>() + right.bytes;
    for row in 0..paths.len() {
        if row % 256 == 0 {
            ctx.memory(
                source_bytes + (next.len() + paths.len()) as u64 * next.rows.len() as u64 * 16,
                "query-hash-join-probe",
            )?;
        }
        let value = path_value(paths, sources, row, left);
        let matching = if matches!(value, Scalar::Null) {
            None
        } else {
            index.get(&scalar_key(&value))
        };
        let mut append = |right_row: Option<usize>| {
            for source in 0..paths.rows.len() {
                next.rows[source].push(paths.rows[source][row]);
            }
            next.rows[source_id].push(right_row);
        };
        if let Some(matching) = matching {
            for right_row in matching {
                append(Some(*right_row));
            }
        } else if matches!(join.kind, JoinKind::Left) {
            append(None);
        }
    }
    *paths = next;
    sources.push(right);
    bindings.extend(new_bindings);
    Ok(())
}
fn grouped_query(
    ctx: &TaskContext<'_>,
    r: &QueryRequest,
    paths: &RowPaths,
    sources: &[Arc<SourceIndex>],
    bindings: &HashMap<u32, ColumnBinding>,
    source_bytes: u64,
) -> KernelResult<QueryResult> {
    let group_bindings = r
        .group_by
        .iter()
        .map(|column| binding(bindings, *column))
        .collect::<KernelResult<Vec<_>>>()?;
    let aggregate_bindings = r
        .aggregates
        .iter()
        .map(|aggregate| binding(bindings, aggregate.column))
        .collect::<KernelResult<Vec<_>>>()?;
    let mut result_columns = r.group_by.clone();
    let mut unique = result_columns.iter().copied().collect::<HashSet<_>>();
    for aggregate in &r.aggregates {
        if !unique.insert(aggregate.output_column) {
            return Err(KernelError::new(
                "QUERY_COLUMN_COLLISION",
                "Aggregate output columns must be unique and distinct from group keys",
            ));
        }
        result_columns.push(aggregate.output_column);
    }
    let columns = r.columns.clone().unwrap_or_else(|| result_columns.clone());
    let ordinal = |column: u32| {
        result_columns
            .iter()
            .position(|candidate| *candidate == column)
            .ok_or_else(|| {
                KernelError::new(
                    "QUERY_GROUP_PROJECTION_INVALID",
                    "Grouped projection may reference only group keys or aggregate outputs",
                )
                .at(column.to_string())
            })
    };
    let selected = columns
        .iter()
        .map(|column| ordinal(*column))
        .collect::<KernelResult<Vec<_>>>()?;
    let sorted = r
        .sort
        .iter()
        .map(|sort| ordinal(sort.column))
        .collect::<KernelResult<Vec<_>>>()?;
    let mut groups = Vec::<GroupRow>::new();
    let mut ids = HashMap::<Vec<String>, usize>::new();
    let mut bytes = source_bytes;
    for row in 0..paths.len() {
        if row % 256 == 0 {
            ctx.memory(bytes, "query-hash-group")?;
        }
        let keys = group_bindings
            .iter()
            .map(|binding| path_value(paths, sources, row, *binding))
            .collect::<Vec<_>>();
        let codes = keys.iter().map(scalar_key).collect::<Vec<_>>();
        let id = if let Some(id) = ids.get(&codes) {
            *id
        } else {
            let id = groups.len();
            bytes += codes.iter().map(|key| key.len() as u64 + 40).sum::<u64>()
                + r.aggregates.len() as u64 * 160;
            ids.insert(codes, id);
            groups.push(GroupRow {
                keys,
                states: r
                    .aggregates
                    .iter()
                    .map(|aggregate| AggregateState::new(aggregate.aggregate))
                    .collect(),
            });
            id
        };
        for (state, binding) in groups[id].states.iter_mut().zip(&aggregate_bindings) {
            let before = state.bytes();
            state.add(&path_value(paths, sources, row, *binding));
            bytes += state.bytes() - before;
        }
    }
    if paths.len() == 0 && r.group_by.is_empty() {
        groups.push(GroupRow {
            keys: vec![],
            states: r
                .aggregates
                .iter()
                .map(|aggregate| AggregateState::new(aggregate.aggregate))
                .collect(),
        });
    }
    let value = |row: &GroupRow, ordinal: usize| {
        if ordinal < row.keys.len() {
            row.keys[ordinal].clone()
        } else {
            let index = ordinal - row.keys.len();
            row.states[index].value(r.aggregates[index].aggregate)
        }
    };
    let mut order = (0..groups.len()).collect::<Vec<_>>();
    order.sort_by(|left, right| {
        for (sort, ordinal) in r.sort.iter().zip(&sorted) {
            let cmp = compare_scalar(
                &value(&groups[*left], *ordinal),
                &value(&groups[*right], *ordinal),
            );
            if !cmp.is_eq() {
                return if sort.descending { cmp.reverse() } else { cmp };
            }
        }
        left.cmp(right)
    });
    ctx.memory(bytes, "query-group-publication")?;
    let total = groups.len();
    let rows = order
        .into_iter()
        .skip(r.offset)
        .take(r.limit)
        .map(|id| QueryRow {
            row_id: id as u32,
            source_row_ids: vec![],
            values: selected
                .iter()
                .map(|ordinal| value(&groups[id], *ordinal))
                .collect(),
        })
        .collect();
    Ok(QueryResult {
        revision: r.revision,
        columns,
        rows,
        total,
        grouped: true,
    })
}
