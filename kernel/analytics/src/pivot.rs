pub use crate::aggregate::AggregateKind;
use crate::{
    aggregate::{AggregateState, compare_scalar, scalar_key},
    source::{SourceColumn, SourceIndex},
    task::TaskContext,
};
use kernel_core::{KernelError, KernelResult, RangeRef, Scalar};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeSet, HashMap, HashSet};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PivotRequest {
    pub revision: u64,
    pub source: RangeRef,
    #[serde(default)]
    pub row_fields: Vec<AxisField>,
    #[serde(default)]
    pub column_fields: Vec<AxisField>,
    #[serde(default)]
    pub value_fields: Vec<ValueField>,
    #[serde(default)]
    pub filters: Vec<crate::FilterColumn>,
    #[serde(default)]
    pub value_filters: Vec<ValueFilter>,
    #[serde(default)]
    pub include_row_totals: bool,
    #[serde(default)]
    pub include_column_totals: bool,
    #[serde(default)]
    pub include_subtotals: bool,
    #[serde(default)]
    pub viewport: Option<Viewport>,
    #[serde(default)]
    pub drilldown: Option<DrilldownRequest>,
    #[serde(default)]
    pub budget: Option<kernel_core::TaskBudget>,
    #[serde(default)]
    pub now_serial: Option<f64>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct AxisField {
    pub column: u32,
    #[serde(default)]
    pub group: Option<crate::pivot_group::Group>,
    #[serde(default)]
    pub sort: Option<AxisSort>,
    #[serde(default)]
    pub subtotal: Option<Subtotal>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "by", rename_all = "camelCase", deny_unknown_fields)]
pub enum AxisSort {
    Label {
        direction: Direction,
    },
    Value {
        direction: Direction,
        #[serde(rename = "valueId")]
        value_id: String,
    },
}
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Direction {
    Ascending,
    Descending,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "mode", rename_all = "camelCase", deny_unknown_fields)]
pub enum Subtotal {
    Automatic,
    None,
    Custom { functions: Vec<AggregateKind> },
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ValueField {
    pub value_id: String,
    pub column: u32,
    pub aggregate: AggregateKind,
    #[serde(default)]
    pub name: Option<String>,
    #[serde(default)]
    pub show_as: Option<ShowAs>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "kebab-case", deny_unknown_fields)]
pub enum ShowAs {
    Normal,
    GrandPercentage,
    RowPercentage,
    ColumnPercentage,
    ParentPercentage,
    Index,
    Difference {
        #[serde(rename = "baseColumn")]
        base_column: u32,
        #[serde(rename = "baseItem")]
        base_item: BaseItem,
    },
    PercentageDifference {
        #[serde(rename = "baseColumn")]
        base_column: u32,
        #[serde(rename = "baseItem")]
        base_item: BaseItem,
    },
    RunningTotal {
        #[serde(rename = "baseColumn")]
        base_column: u32,
    },
    PercentageRunningTotal {
        #[serde(rename = "baseColumn")]
        base_column: u32,
    },
    Rank {
        #[serde(rename = "baseColumn")]
        base_column: u32,
        direction: Direction,
    },
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum BaseItem {
    Relative(RelativeItem),
    Member(crate::pivot_group::MemberKey),
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum RelativeItem {
    Previous,
    Next,
}
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Axis {
    Rows,
    Columns,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ValueFilter {
    pub axis: Axis,
    pub depth: usize,
    pub value_id: String,
    #[serde(default)]
    pub predicate: Option<crate::FilterPredicate>,
    #[serde(default)]
    pub top: Option<TopFilter>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TopFilter {
    pub direction: TopDirection,
    pub mode: TopMode,
    pub threshold: f64,
}
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TopDirection {
    Top,
    Bottom,
}
#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TopMode {
    Items,
    Percent,
    Sum,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Viewport {
    pub row_offset: usize,
    pub column_offset: usize,
    pub row_limit: usize,
    pub column_limit: usize,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct DrilldownRequest {
    pub row_keys: Vec<Scalar>,
    pub column_keys: Vec<Scalar>,
    pub offset: usize,
    pub limit: usize,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotResult {
    pub revision: u64,
    pub rows: Vec<PivotAxisNode>,
    pub columns: Vec<PivotColumnNode>,
    pub cells: Vec<PivotCell>,
    pub fields: Vec<ValueField>,
    pub total_groups: usize,
    pub total_columns: usize,
    pub source_rows: usize,
    pub drilldown: Option<DrilldownPage>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotAxisNode {
    pub row_id: usize,
    pub keys: Vec<Scalar>,
    pub subtotal: bool,
    pub grand_total: bool,
    pub subtotal_function: Option<AggregateKind>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotColumnNode {
    pub column_id: usize,
    pub keys: Vec<Scalar>,
    pub subtotal: bool,
    pub grand_total: bool,
    pub subtotal_function: Option<AggregateKind>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PivotCell {
    pub row_id: usize,
    pub column_id: usize,
    pub values: Vec<Scalar>,
}
#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DrilldownPage {
    pub source_rows: Vec<u32>,
    pub values: Vec<Vec<Scalar>>,
    pub columns: Vec<u32>,
    pub total: usize,
    pub offset: usize,
}

struct AxisVector {
    members: Vec<Scalar>,
    codes: Vec<u32>,
}
#[derive(Default)]
struct PathPool {
    paths: Vec<Vec<u32>>,
    ids: HashMap<Vec<u32>, usize>,
}
impl PathPool {
    fn intern(&mut self, path: &[u32]) -> usize {
        if let Some(id) = self.ids.get(path) {
            return *id;
        }
        let id = self.paths.len();
        self.paths.push(path.to_vec());
        self.ids.insert(path.to_vec(), id);
        id
    }
}
type States = Vec<AggregateState>;
type SparseStates = HashMap<(usize, usize), States>;
type Rollups = HashMap<(Vec<u32>, Vec<u32>), States>;
struct ResultNode {
    path: Vec<u32>,
    function: Option<AggregateKind>,
}

pub(crate) fn execute_pivot(
    ctx: &TaskContext<'_>,
    request: PivotRequest,
) -> KernelResult<PivotResult> {
    validate(&request)?;
    let selected = request
        .row_fields
        .iter()
        .chain(&request.column_fields)
        .map(|f| f.column)
        .chain(request.value_fields.iter().map(|f| f.column))
        .chain(request.filters.iter().map(|f| f.column))
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect::<Vec<_>>();
    let source = SourceIndex::build(ctx, &request.source, &selected)?;
    let filter_masks =
        crate::filter::filter_masks(ctx, &source, &request.filters, request.now_serial)?;
    let row_vectors = axis_vectors(ctx, &source, &request.row_fields)?;
    let column_vectors = axis_vectors(ctx, &source, &request.column_fields)?;
    let mut row_pool = PathPool::default();
    let mut column_pool = PathPool::default();
    let mut grouped = SparseStates::new();
    // One row id pair per input row is enough for lazy drilldown. No per-cell provenance copies.
    let mut row_membership = vec![usize::MAX; source.row_count];
    let mut column_membership = vec![usize::MAX; source.row_count];
    let mut row_key = vec![0; row_vectors.len()];
    let mut column_key = vec![0; column_vectors.len()];
    let mut bytes = source.bytes
        + (source.row_count
            * (16 + 4 * (row_vectors.len() + column_vectors.len()) + request.filters.len()))
            as u64;
    for row in 0..source.row_count {
        if row % 256 == 0 {
            ctx.memory(bytes, "pivot-aggregate")?;
        }
        if !filter_masks.iter().all(|mask| mask[row]) {
            continue;
        }
        for (key, vector) in row_key.iter_mut().zip(&row_vectors) {
            *key = vector.codes[row];
        }
        for (key, vector) in column_key.iter_mut().zip(&column_vectors) {
            *key = vector.codes[row];
        }
        let row_id = row_pool.intern(&row_key);
        let column_id = column_pool.intern(&column_key);
        row_membership[row] = row_id;
        column_membership[row] = column_id;
        let states = grouped.entry((row_id, column_id)).or_insert_with(|| {
            bytes += new_state_bytes(&request);
            new_states(&request)
        });
        for (state, field) in states.iter_mut().zip(&request.value_fields) {
            let before = state.bytes();
            state.add(&source.value_at(field.column, row)?);
            bytes += state.bytes() - before;
        }
    }
    if row_pool.paths.is_empty() && request.row_fields.is_empty() {
        row_pool.intern(&[]);
    }
    if column_pool.paths.is_empty() && request.column_fields.is_empty() {
        column_pool.intern(&[]);
    }
    let (allowed_rows, allowed_columns) =
        filter_axes(ctx, &request, &grouped, &row_pool, &column_pool)?;
    grouped
        .retain(|(row, column), _| allowed_rows.contains(row) && allowed_columns.contains(column));
    let row_nodes = make_nodes(
        &request.row_fields,
        &row_pool,
        &allowed_rows,
        &request,
        request.include_column_totals,
    );
    let column_nodes = make_nodes(
        &request.column_fields,
        &column_pool,
        &allowed_columns,
        &request,
        request.include_row_totals,
    );
    let (row_nodes, column_nodes, sparse, rollups) = physical_pivot(
        ctx,
        &request,
        &grouped,
        &row_pool,
        &column_pool,
        row_nodes,
        column_nodes,
        &row_vectors,
        &column_vectors,
        bytes,
    )?;
    let total_groups = row_nodes.len();
    let total_columns = column_nodes.len();
    let viewport = request.viewport.clone().unwrap_or(Viewport {
        row_offset: 0,
        column_offset: 0,
        row_limit: total_groups,
        column_limit: total_columns,
    });
    let row_end = viewport
        .row_offset
        .saturating_add(viewport.row_limit)
        .min(total_groups);
    let column_end = viewport
        .column_offset
        .saturating_add(viewport.column_limit)
        .min(total_columns);
    let mut cells = Vec::new();
    for (&(row, column), states) in &sparse {
        if row < viewport.row_offset
            || row >= row_end
            || column < viewport.column_offset
            || column >= column_end
        {
            continue;
        }
        if cells.len() % 256 == 0 {
            ctx.checkpoint("pivot-viewport")?;
        }
        let values = request
            .value_fields
            .iter()
            .enumerate()
            .map(|(i, field)| {
                let value = states[i].value(
                    row_nodes[row]
                        .function
                        .or(column_nodes[column].function)
                        .unwrap_or(field.aggregate),
                );
                transform(
                    value,
                    i,
                    row,
                    column,
                    &request,
                    &row_nodes,
                    &column_nodes,
                    &sparse,
                    &rollups,
                    &row_vectors,
                    &column_vectors,
                )
            })
            .collect();
        cells.push(PivotCell {
            row_id: row,
            column_id: column,
            values,
        });
    }
    cells.sort_by_key(|cell| (cell.row_id, cell.column_id));
    let drilldown = request
        .drilldown
        .as_ref()
        .map(|drill| {
            drilldown_page(
                ctx,
                drill,
                &source,
                &row_pool,
                &column_pool,
                &row_vectors,
                &column_vectors,
                &row_membership,
                &column_membership,
                &allowed_rows,
                &allowed_columns,
            )
        })
        .transpose()?;
    let rows = row_nodes
        .iter()
        .enumerate()
        .skip(viewport.row_offset)
        .take(viewport.row_limit)
        .map(|(id, node)| axis_node(id, node, &row_vectors))
        .collect();
    let columns = column_nodes
        .iter()
        .enumerate()
        .skip(viewport.column_offset)
        .take(viewport.column_limit)
        .map(|(id, node)| {
            let axis = axis_node(id, node, &column_vectors);
            PivotColumnNode {
                column_id: id,
                keys: axis.keys,
                subtotal: axis.subtotal,
                grand_total: axis.grand_total,
                subtotal_function: axis.subtotal_function,
            }
        })
        .collect();
    Ok(PivotResult {
        revision: request.revision,
        rows,
        columns,
        cells,
        fields: request.value_fields,
        total_groups,
        total_columns,
        source_rows: source.row_count,
        drilldown,
    })
}

fn validate(r: &PivotRequest) -> KernelResult<()> {
    r.source.validate()?;
    if r.value_fields.is_empty() {
        return Err(KernelError::new(
            "PIVOT_VALUES_REQUIRED",
            "Pivot requires at least one Values placement",
        ));
    }
    let mut ids = HashSet::new();
    for field in &r.value_fields {
        if field.value_id.is_empty() || !ids.insert(&field.value_id) {
            return Err(KernelError::new(
                "PIVOT_VALUE_ID_INVALID",
                "Values placement identities must be nonempty and unique",
            ));
        }
    }
    for field in r.row_fields.iter().chain(&r.column_fields) {
        if let Some(AxisSort::Value { value_id, .. }) = &field.sort {
            value_index(r, value_id)?;
        }
        if let Some(Subtotal::Custom { functions }) = &field.subtotal {
            if functions.is_empty() {
                return Err(KernelError::new(
                    "PIVOT_SUBTOTAL_INVALID",
                    "Custom subtotals require at least one function",
                ));
            }
        }
    }
    for field in &r.value_fields {
        if let Some(spec) = &field.show_as {
            if let Some(base) = base_column(spec) {
                if !r
                    .row_fields
                    .iter()
                    .chain(&r.column_fields)
                    .any(|axis| axis.column == base)
                {
                    return Err(KernelError::new(
                        "PIVOT_SHOW_AS_BASE_INVALID",
                        "Show Values As base field must be on a row or column axis",
                    ));
                }
            }
        }
    }
    for filter in &r.value_filters {
        value_index(r, &filter.value_id)?;
        let axes = match filter.axis {
            Axis::Rows => &r.row_fields,
            Axis::Columns => &r.column_fields,
        };
        if filter.depth >= axes.len() || filter.predicate.is_some() == filter.top.is_some() {
            return Err(KernelError::new(
                "PIVOT_VALUE_FILTER_INVALID",
                "Value filter requires a valid axis depth and exactly one predicate or top filter",
            ));
        }
        if let Some(top) = &filter.top {
            if !top.threshold.is_finite()
                || top.threshold <= 0.
                || matches!(top.mode, TopMode::Percent) && top.threshold > 100.
                || matches!(top.mode, TopMode::Items) && top.threshold.fract() != 0.
            {
                return Err(KernelError::new(
                    "PIVOT_TOP_FILTER_INVALID",
                    "Top filter threshold is invalid for its mode",
                ));
            }
        }
    }
    Ok(())
}
fn value_index(r: &PivotRequest, id: &str) -> KernelResult<usize> {
    r.value_fields
        .iter()
        .position(|f| f.value_id == id)
        .ok_or_else(|| {
            KernelError::new(
                "PIVOT_VALUE_FIELD_NOT_FOUND",
                format!("Values placement {id} does not exist"),
            )
        })
}
fn new_states(r: &PivotRequest) -> States {
    let distinct_subtotal=r.row_fields.iter().chain(&r.column_fields).any(|f|matches!(&f.subtotal,Some(Subtotal::Custom{functions}) if functions.contains(&AggregateKind::DistinctCount)));
    r.value_fields
        .iter()
        .map(|f| {
            AggregateState::new(if distinct_subtotal {
                AggregateKind::DistinctCount
            } else {
                f.aggregate
            })
        })
        .collect()
}
fn new_state_bytes(r: &PivotRequest) -> u64 {
    new_states(r).iter().map(AggregateState::bytes).sum::<u64>() + 64
}
fn merge(target: &mut States, source: &States) {
    for (a, b) in target.iter_mut().zip(source) {
        a.merge(b);
    }
}
fn axis_vectors(
    ctx: &TaskContext<'_>,
    source: &SourceIndex,
    fields: &[AxisField],
) -> KernelResult<Vec<AxisVector>> {
    fields
        .iter()
        .map(|field| {
            let mut members = vec![Scalar::Null];
            let mut ids = HashMap::from([(scalar_key(&Scalar::Null), 0u32)]);
            let mut intern = |value: Scalar| -> KernelResult<u32> {
                let value = if let Some(group) = &field.group {
                    crate::pivot_group::group(&value, group)?
                } else {
                    value
                };
                let key = scalar_key(&value);
                if let Some(id) = ids.get(&key) {
                    return Ok(*id);
                }
                let id = members.len() as u32;
                members.push(value);
                ids.insert(key, id);
                Ok(id)
            };
            let mut codes = vec![0; source.row_count];
            match source.column(field.column)? {
                SourceColumn::Dictionary {
                    dictionary,
                    codes: source_codes,
                } => {
                    let mut mapped = vec![0];
                    for value in dictionary {
                        mapped.push(intern(value.clone())?);
                    }
                    for (row, code) in source_codes.iter().enumerate() {
                        codes[row] = mapped[*code as usize];
                    }
                }
                _ => {
                    for (row, code) in codes.iter_mut().enumerate() {
                        if row % 1024 == 0 {
                            ctx.checkpoint("pivot-axis-index")?;
                        }
                        *code = intern(source.value_at(field.column, row)?)?;
                    }
                }
            }
            Ok(AxisVector { members, codes })
        })
        .collect()
}
fn filter_axes(
    ctx: &TaskContext<'_>,
    r: &PivotRequest,
    grouped: &SparseStates,
    rows: &PathPool,
    columns: &PathPool,
) -> KernelResult<(HashSet<usize>, HashSet<usize>)> {
    let mut allowed_rows = (0..rows.paths.len()).collect::<HashSet<_>>();
    let mut allowed_columns = (0..columns.paths.len()).collect::<HashSet<_>>();
    for filter in &r.value_filters {
        ctx.checkpoint("pivot-value-filter")?;
        let index = value_index(r, &filter.value_id)?;
        let pool = match filter.axis {
            Axis::Rows => rows,
            Axis::Columns => columns,
        };
        let mut totals = HashMap::<Vec<u32>, States>::new();
        for (&(row, column), states) in grouped {
            if !allowed_rows.contains(&row) || !allowed_columns.contains(&column) {
                continue;
            }
            let id = match filter.axis {
                Axis::Rows => row,
                Axis::Columns => column,
            };
            let path = &pool.paths[id][..=filter.depth];
            merge(
                totals.entry(path.to_vec()).or_insert_with(|| new_states(r)),
                states,
            );
        }
        let allowed = if let Some(predicate) = &filter.predicate {
            totals
                .iter()
                .filter(|(_, states)| {
                    crate::filter::matches_pred(
                        &states[index].value(r.value_fields[index].aggregate),
                        predicate,
                    )
                })
                .map(|(key, _)| key.clone())
                .collect::<HashSet<_>>()
        } else {
            let top = filter.top.as_ref().expect("validated top filter");
            let mut siblings = HashMap::<Vec<u32>, Vec<(Vec<u32>, f64)>>::new();
            for (path, states) in &totals {
                if let Scalar::Number(score) = states[index].value(r.value_fields[index].aggregate)
                {
                    siblings
                        .entry(path[..filter.depth].to_vec())
                        .or_default()
                        .push((path.clone(), score));
                }
            }
            let mut allowed = HashSet::new();
            for scores in siblings.values_mut() {
                scores.sort_by(|a, b| {
                    let order = a.1.total_cmp(&b.1);
                    (if matches!(top.direction, TopDirection::Top) {
                        order.reverse()
                    } else {
                        order
                    })
                    .then(a.0.cmp(&b.0))
                });
                let count = match top.mode {
                    TopMode::Items => top.threshold as usize,
                    TopMode::Percent => {
                        (scores.len() as f64 * top.threshold / 100.).ceil() as usize
                    }
                    TopMode::Sum => scores.len(),
                };
                let mut sum = 0.;
                let mut cutoff = None;
                for (position, (path, score)) in scores.iter().enumerate() {
                    let include = match top.mode {
                        TopMode::Sum => sum < top.threshold,
                        _ => position < count || cutoff == Some(*score),
                    };
                    if !include {
                        break;
                    }
                    allowed.insert(path.clone());
                    sum += score;
                    if position + 1 == count {
                        cutoff = Some(*score);
                    }
                }
            }
            allowed
        };
        let target = match filter.axis {
            Axis::Rows => &mut allowed_rows,
            Axis::Columns => &mut allowed_columns,
        };
        target.retain(|id| allowed.contains(&pool.paths[*id][..=filter.depth]));
    }
    Ok((allowed_rows, allowed_columns))
}
fn make_nodes(
    fields: &[AxisField],
    pool: &PathPool,
    allowed: &HashSet<usize>,
    r: &PivotRequest,
    include_totals: bool,
) -> Vec<ResultNode> {
    let mut paths = BTreeSet::new();
    for id in allowed {
        paths.insert(pool.paths[*id].clone());
        if r.include_subtotals {
            for depth in 1..pool.paths[*id].len() {
                if !matches!(fields[depth - 1].subtotal, Some(Subtotal::None)) {
                    paths.insert(pool.paths[*id][..depth].to_vec());
                }
            }
        }
    }
    if include_totals || fields.is_empty() {
        paths.insert(Vec::new());
    }
    let mut nodes = Vec::new();
    for path in paths {
        if !path.is_empty() && path.len() < fields.len() {
            if let Some(Subtotal::Custom { functions }) = &fields[path.len() - 1].subtotal {
                for function in functions {
                    nodes.push(ResultNode {
                        path: path.clone(),
                        function: Some(*function),
                    });
                }
                continue;
            }
        }
        nodes.push(ResultNode {
            path,
            function: None,
        });
    }
    nodes
}

#[allow(clippy::too_many_arguments)]
fn physical_pivot(
    ctx: &TaskContext<'_>,
    r: &PivotRequest,
    grouped: &SparseStates,
    row_pool: &PathPool,
    column_pool: &PathPool,
    mut rows: Vec<ResultNode>,
    mut columns: Vec<ResultNode>,
    row_vectors: &[AxisVector],
    column_vectors: &[AxisVector],
    mut bytes: u64,
) -> KernelResult<(Vec<ResultNode>, Vec<ResultNode>, SparseStates, Rollups)> {
    // Ancestor states are merged from each occupied aggregate once. Their domain
    // is the actual sparse group tree, never the row/column Cartesian product.
    let mut rollups = Rollups::new();
    for (position, (&(row, column), states)) in grouped.iter().enumerate() {
        if position % 256 == 0 {
            ctx.memory(bytes, "pivot-rollup")?;
        }
        for rd in 0..=row_pool.paths[row].len() {
            for cd in 0..=column_pool.paths[column].len() {
                let target = rollups
                    .entry((
                        row_pool.paths[row][..rd].to_vec(),
                        column_pool.paths[column][..cd].to_vec(),
                    ))
                    .or_insert_with(|| {
                        bytes += new_state_bytes(r);
                        new_states(r)
                    });
                for (a, b) in target.iter_mut().zip(states) {
                    let before = a.bytes();
                    a.merge(b);
                    bytes += a.bytes() - before;
                }
            }
        }
    }
    sort_nodes(
        &mut rows,
        &r.row_fields,
        row_vectors,
        &rollups,
        r,
        Axis::Rows,
    );
    sort_nodes(
        &mut columns,
        &r.column_fields,
        column_vectors,
        &rollups,
        r,
        Axis::Columns,
    );
    let mut row_targets = HashMap::<Vec<u32>, Vec<usize>>::new();
    let mut column_targets = HashMap::<Vec<u32>, Vec<usize>>::new();
    for (id, node) in rows.iter().enumerate() {
        row_targets.entry(node.path.clone()).or_default().push(id);
    }
    for (id, node) in columns.iter().enumerate() {
        column_targets
            .entry(node.path.clone())
            .or_default()
            .push(id);
    }
    let mut result = SparseStates::new();
    for ((row_path, column_path), states) in &rollups {
        if let (Some(row_ids), Some(column_ids)) =
            (row_targets.get(row_path), column_targets.get(column_path))
        {
            for row_id in row_ids {
                for column_id in column_ids {
                    bytes += states.iter().map(AggregateState::bytes).sum::<u64>();
                    result.insert((*row_id, *column_id), states.clone());
                }
            }
        }
    }
    ctx.memory(bytes, "pivot-physical-publish")?;
    Ok((rows, columns, result, rollups))
}
fn sort_nodes(
    nodes: &mut [ResultNode],
    fields: &[AxisField],
    vectors: &[AxisVector],
    totals: &Rollups,
    r: &PivotRequest,
    axis: Axis,
) {
    nodes.sort_by(|a, b| {
        if a.path.is_empty() != b.path.is_empty() {
            return if a.path.is_empty() {
                std::cmp::Ordering::Greater
            } else {
                std::cmp::Ordering::Less
            };
        }
        for depth in 0..a.path.len().min(b.path.len()) {
            if a.path[depth] == b.path[depth] {
                continue;
            }
            let field = &fields[depth];
            let label = compare_scalar(
                &vectors[depth].members[a.path[depth] as usize],
                &vectors[depth].members[b.path[depth] as usize],
            );
            let (order, direction) = match &field.sort {
                Some(AxisSort::Value {
                    value_id,
                    direction,
                }) => {
                    let index = r
                        .value_fields
                        .iter()
                        .position(|v| v.value_id == *value_id)
                        .expect("validated value field");
                    let key = |path: &[u32]| match axis {
                        Axis::Rows => (path.to_vec(), vec![]),
                        Axis::Columns => (vec![], path.to_vec()),
                    };
                    let left = totals
                        .get(&key(&a.path[..=depth]))
                        .map(|s| s[index].value(r.value_fields[index].aggregate))
                        .unwrap_or(Scalar::Null);
                    let right = totals
                        .get(&key(&b.path[..=depth]))
                        .map(|s| s[index].value(r.value_fields[index].aggregate))
                        .unwrap_or(Scalar::Null);
                    (compare_scalar(&left, &right).then(label), *direction)
                }
                Some(AxisSort::Label { direction }) => (label, *direction),
                None => (label, Direction::Ascending),
            };
            return if matches!(direction, Direction::Descending) {
                order.reverse()
            } else {
                order
            };
        }
        b.path.len().cmp(&a.path.len())
    })
}
fn axis_node(id: usize, node: &ResultNode, vectors: &[AxisVector]) -> PivotAxisNode {
    PivotAxisNode {
        row_id: id,
        keys: node
            .path
            .iter()
            .enumerate()
            .map(|(i, code)| vectors[i].members[*code as usize].clone())
            .collect(),
        subtotal: !node.path.is_empty() && node.path.len() < vectors.len(),
        grand_total: node.path.is_empty() && !vectors.is_empty(),
        subtotal_function: node.function,
    }
}
fn numeric(v: Scalar) -> Option<f64> {
    if let Scalar::Number(n) = v {
        Some(n)
    } else {
        None
    }
}
fn ratio(a: f64, b: Option<f64>) -> Scalar {
    match b {
        Some(b) if b != 0. => Scalar::Number(a / b),
        _ => Scalar::Null,
    }
}
fn base_column(show: &ShowAs) -> Option<u32> {
    match show {
        ShowAs::Difference { base_column, .. }
        | ShowAs::PercentageDifference { base_column, .. }
        | ShowAs::RunningTotal { base_column }
        | ShowAs::PercentageRunningTotal { base_column }
        | ShowAs::Rank { base_column, .. } => Some(*base_column),
        _ => None,
    }
}
#[allow(clippy::too_many_arguments)]
fn transform(
    value: Scalar,
    index: usize,
    row: usize,
    column: usize,
    r: &PivotRequest,
    rows: &[ResultNode],
    columns: &[ResultNode],
    sparse: &SparseStates,
    rollups: &Rollups,
    row_vectors: &[AxisVector],
    column_vectors: &[AxisVector],
) -> Scalar {
    let Some(show) = &r.value_fields[index].show_as else {
        return value;
    };
    let Scalar::Number(current) = value else {
        return value;
    };
    let aggregate = r.value_fields[index].aggregate;
    let read = |ri: usize, ci: usize| {
        sparse
            .get(&(ri, ci))
            .and_then(|states| numeric(states[index].value(aggregate)))
    };
    let total = |rp: &[u32], cp: &[u32]| {
        rollups
            .get(&(rp.to_vec(), cp.to_vec()))
            .and_then(|states| numeric(states[index].value(aggregate)))
    };
    let grand = total(&[], &[]);
    let row_total = total(&rows[row].path, &[]);
    let column_total = total(&[], &columns[column].path);
    match show {
        ShowAs::Normal => Scalar::Number(current),
        ShowAs::GrandPercentage => ratio(current, grand),
        ShowAs::RowPercentage => ratio(current, row_total),
        ShowAs::ColumnPercentage => ratio(current, column_total),
        ShowAs::ParentPercentage => ratio(
            current,
            total(
                &rows[row].path[..rows[row].path.len().saturating_sub(1)],
                &columns[column].path,
            ),
        ),
        ShowAs::Index => match (grand, row_total, column_total) {
            (Some(g), Some(rt), Some(ct)) if rt != 0. && ct != 0. => {
                Scalar::Number(current * g / rt / ct)
            }
            _ => Scalar::Null,
        },
        _ => {
            let base = base_column(show).expect("axis mode");
            let (axis, nodes, fields, vectors, position) =
                if r.row_fields.iter().any(|f| f.column == base) {
                    (Axis::Rows, rows, &r.row_fields, row_vectors, row)
                } else {
                    (
                        Axis::Columns,
                        columns,
                        &r.column_fields,
                        column_vectors,
                        column,
                    )
                };
            let depth = fields
                .iter()
                .position(|f| f.column == base)
                .expect("validated base");
            let path = &nodes[position].path;
            if path.len() <= depth {
                return Scalar::Null;
            }
            let series = nodes
                .iter()
                .enumerate()
                .filter(|(_, node)| {
                    node.path.len() == path.len()
                        && node.function == nodes[position].function
                        && node
                            .path
                            .iter()
                            .enumerate()
                            .all(|(i, v)| i == depth || *v == path[i])
                })
                .map(|(id, _)| {
                    (
                        id,
                        match axis {
                            Axis::Rows => read(id, column),
                            Axis::Columns => read(row, id),
                        },
                    )
                })
                .collect::<Vec<_>>();
            let Some(current_position) = series.iter().position(|(id, _)| *id == position) else {
                return Scalar::Null;
            };
            match show {
                ShowAs::RunningTotal { .. } | ShowAs::PercentageRunningTotal { .. } => {
                    let cumulative = series[..=current_position]
                        .iter()
                        .filter_map(|(_, value)| *value)
                        .sum::<f64>();
                    if matches!(show, ShowAs::RunningTotal { .. }) {
                        Scalar::Number(cumulative)
                    } else {
                        ratio(
                            cumulative,
                            Some(series.iter().filter_map(|(_, value)| *value).sum()),
                        )
                    }
                }
                ShowAs::Rank { direction, .. } => Scalar::Number(
                    1. + series
                        .iter()
                        .filter_map(|(_, value)| *value)
                        .filter(|value| {
                            if matches!(direction, Direction::Ascending) {
                                *value < current
                            } else {
                                *value > current
                            }
                        })
                        .count() as f64,
                ),
                ShowAs::Difference { base_item, .. }
                | ShowAs::PercentageDifference { base_item, .. } => {
                    let target = match base_item {
                        BaseItem::Relative(RelativeItem::Previous) => {
                            current_position.checked_sub(1)
                        }
                        BaseItem::Relative(RelativeItem::Next) => {
                            (current_position + 1 < series.len()).then_some(current_position + 1)
                        }
                        BaseItem::Member(member) => series.iter().position(|(id, _)| {
                            member_matches(
                                member,
                                &vectors[depth].members[nodes[*id].path[depth] as usize],
                            )
                        }),
                    };
                    let Some(target) = target.filter(|target| *target != current_position) else {
                        return Scalar::Null;
                    };
                    match series[target].1 {
                        Some(base) => {
                            if matches!(show, ShowAs::Difference { .. }) {
                                Scalar::Number(current - base)
                            } else {
                                ratio(current - base, Some(base))
                            }
                        }
                        None => Scalar::Null,
                    }
                }
                _ => Scalar::Null,
            }
        }
    }
}
fn member_matches(member: &crate::pivot_group::MemberKey, value: &Scalar) -> bool {
    serde_json::to_value(member)
        .ok()
        .and_then(|member| member.get("value").cloned())
        .is_some_and(|member| {
            serde_json::to_value(value)
                .ok()
                .is_some_and(|value| value == member)
        })
}
#[allow(clippy::too_many_arguments)]
fn drilldown_page(
    ctx: &TaskContext<'_>,
    request: &DrilldownRequest,
    source: &SourceIndex,
    row_pool: &PathPool,
    column_pool: &PathPool,
    row_vectors: &[AxisVector],
    column_vectors: &[AxisVector],
    row_membership: &[usize],
    column_membership: &[usize],
    allowed_rows: &HashSet<usize>,
    allowed_columns: &HashSet<usize>,
) -> KernelResult<DrilldownPage> {
    if request.row_keys.len() > row_vectors.len()
        || request.column_keys.len() > column_vectors.len()
    {
        return Err(KernelError::new(
            "PIVOT_DRILLDOWN_COORDINATE_INVALID",
            "Drilldown address has too many axis keys",
        ));
    }
    let matches = |path: &[u32], vectors: &[AxisVector], keys: &[Scalar]| {
        keys.iter().enumerate().all(|(depth, value)| {
            scalar_key(value) == scalar_key(&vectors[depth].members[path[depth] as usize])
        })
    };
    let mut source_rows = Vec::new();
    let mut values = Vec::new();
    let mut total = 0;
    for row in 0..source.row_count {
        if row % 256 == 0 {
            ctx.checkpoint("pivot-drilldown")?;
        }
        let ri = row_membership[row];
        let ci = column_membership[row];
        if !allowed_rows.contains(&ri) || !allowed_columns.contains(&ci) {
            continue;
        }
        if !matches(&row_pool.paths[ri], row_vectors, &request.row_keys)
            || !matches(&column_pool.paths[ci], column_vectors, &request.column_keys)
        {
            continue;
        }
        if total >= request.offset && source_rows.len() < request.limit {
            source_rows.push(source.range.start_row + row as u32);
            values.push(
                source
                    .columns
                    .iter()
                    .map(|column| source.value_at(*column, row))
                    .collect::<KernelResult<Vec<_>>>()?,
            );
        }
        total += 1;
    }
    Ok(DrilldownPage {
        source_rows,
        values,
        columns: source.columns.clone(),
        total,
        offset: request.offset,
    })
}
