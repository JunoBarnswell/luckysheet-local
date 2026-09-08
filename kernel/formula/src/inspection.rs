//! Read-only formula inspection and controlled evaluation surfaces.
//!
//! The methods here deliberately use the same parser, dependency index and
//! evaluator as normal calculation. They do not maintain a second formula
//! graph or repair values in the inspection layer.

use super::*;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CellDependency {
    pub kind: String,
    pub address: CellAddress,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RangeDependency {
    pub kind: String,
    pub start: CellAddress,
    pub end: CellAddress,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NameDependency {
    pub kind: String,
    pub name: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(untagged)]
pub enum FormulaDependency {
    Cell(CellDependency),
    Range(RangeDependency),
    Name(NameDependency),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaCellEntry {
    pub address: CellAddress,
    pub formula: String,
    pub value: FormulaValue,
    pub dependencies: Vec<FormulaDependency>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ResolvedSpill {
    pub sheet_id: String,
    pub anchor: SpillAnchor,
    pub range: RangeRef,
    pub values: Vec<Vec<Scalar>>,
    pub state: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpillAnchor {
    pub row: u32,
    pub column: u32,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaInspection {
    pub revision: u64,
    pub generation: u64,
    pub entries: Vec<FormulaCellEntry>,
    pub dependents: Vec<CellAddress>,
    pub spills: Vec<ResolvedSpill>,
    pub pending_recalculation: bool,
    pub next_cursor: Option<String>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct InspectionQuery {
    pub address: Option<CellAddress>,
    pub projection: Option<String>,
    pub sheet_id: Option<String>,
    pub cursor: Option<String>,
    pub limit: Option<usize>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaEvaluationTraceStep {
    pub expression: String,
    pub value: FormulaValue,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaTrace {
    pub value: FormulaValue,
    pub steps: Vec<FormulaEvaluationTraceStep>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SpillValue {
    pub value: Option<FormulaValue>,
    pub is_spill: bool,
}

fn dependency_values(expr: &Expr) -> Vec<FormulaDependency> {
    let mut out = Vec::new();
    fn walk(expr: &Expr, out: &mut Vec<FormulaDependency>) {
        match expr {
            Expr::Reference(range) => append_range_dependencies(out, std::slice::from_ref(range)),
            Expr::Name(name) => {
                out.push(FormulaDependency::Name(NameDependency {
                    kind: "name".into(),
                    name: name.to_uppercase(),
                }))
            }
            Expr::Structured(reference) => out.push(FormulaDependency::Name(NameDependency {
                kind: "name".into(),
                name: reference.table_name.to_uppercase(),
            })),
            Expr::Array(rows) => {
                for expr in rows.iter().flatten() {
                    walk(expr, out);
                }
            }
            Expr::Unary(_, expr) | Expr::Spill(expr) => walk(expr, out),
            Expr::Binary(_, left, right) => {
                walk(left, out);
                walk(right, out);
            }
            Expr::Call(_, args) | Expr::Invoke(_, args) => {
                for expr in args {
                    walk(expr, out);
                }
            }
            Expr::Scalar(_) | Expr::Missing => {}
        }
    }
    walk(expr, &mut out);
    out
}

fn append_range_dependencies(out: &mut Vec<FormulaDependency>, ranges: &[RangeRef]) {
    for range in ranges {
        if range.start_row == range.end_row && range.start_column == range.end_column {
            out.push(FormulaDependency::Cell(CellDependency {
                kind: "cell".into(),
                address: CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row: range.start_row,
                    column: range.start_column,
                },
            }));
        } else {
            out.push(FormulaDependency::Range(RangeDependency {
                kind: "range".into(),
                start: CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row: range.start_row,
                    column: range.start_column,
                },
                end: CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row: range.end_row,
                    column: range.end_column,
                },
            }));
        }
    }
}

fn dependents(runtime: &FormulaRuntime, target: &CellAddress) -> Vec<CellAddress> {
    let mut result = BTreeSet::new();
    let mut queue = vec![target.clone()];
    while let Some(address) = queue.pop() {
        for dependent in runtime.index.affected(&address) {
            if result.insert(dependent.clone()) {
                queue.push(dependent);
            }
        }
    }
    result.into_iter().collect()
}
fn inspection_bounds(
    query: &InspectionQuery,
) -> KernelResult<(std::ops::Bound<CellAddress>, std::ops::Bound<CellAddress>)> {
    use std::ops::Bound::*;
    if let Some(address) = &query.address {
        address.validate()?;
        return Ok((Included(address.clone()), Included(address.clone())));
    }
    let start = query.sheet_id.as_ref().map(|sheet| CellAddress {
        sheet_id: sheet.clone(),
        row: 0,
        column: 0,
    });
    let end = query.sheet_id.as_ref().map(|sheet| CellAddress {
        sheet_id: sheet.clone(),
        row: MAX_ROWS - 1,
        column: MAX_COLUMNS - 1,
    });
    let lower = if let Some(raw_cursor) = &query.cursor {
        let cursor: CellAddress = serde_json::from_str(raw_cursor).map_err(|_| {
            KernelError::new("INSPECTION_QUERY_INVALID", "Inspection cursor is malformed")
        })?;
        cursor.validate()?;
        if query
            .sheet_id
            .as_ref()
            .is_some_and(|sheet| sheet != &cursor.sheet_id)
        {
            return Err(KernelError::new(
                "INSPECTION_QUERY_INVALID",
                "Inspection cursor belongs to another worksheet",
            ));
        }
        Excluded(cursor)
    } else {
        start.map(Included).unwrap_or(Unbounded)
    };
    Ok((lower, end.map(Included).unwrap_or(Unbounded)))
}

fn spill_projection(address: &CellAddress, value: &FormulaValue) -> Option<ResolvedSpill> {
    let FormulaValue::Array(values) = value else {
        return None;
    };
    let rows = values.len();
    let columns = values.first().map_or(0, Vec::len);
    if rows == 0 || columns == 0 {
        return None;
    }
    Some(ResolvedSpill {
        sheet_id: address.sheet_id.clone(),
        anchor: SpillAnchor {
            row: address.row,
            column: address.column,
        },
        range: RangeRef {
            sheet_id: address.sheet_id.clone(),
            start_row: address.row,
            end_row: address.row + rows as u32 - 1,
            start_column: address.column,
            end_column: address.column + columns as u32 - 1,
        },
        values: values.clone(),
        state: "ok".into(),
    })
}

impl FormulaRuntime {
    pub fn inspect(
        &self,
        reader: &dyn CellReader,
        address: Option<&CellAddress>,
    ) -> KernelResult<FormulaInspection> {
        let mut query = InspectionQuery::default();
        query.address = address.cloned();
        self.inspect_query(reader, &query)
    }

    pub fn inspect_query(
        &self,
        reader: &dyn CellReader,
        query: &InspectionQuery,
    ) -> KernelResult<FormulaInspection> {
        self.inspect_query_with_services(reader, query, &DEFAULT_SERVICES)
    }
    pub fn inspect_query_with_services(
        &self,
        reader: &dyn CellReader,
        query: &InspectionQuery,
        services: &dyn CalculationServices,
    ) -> KernelResult<FormulaInspection> {
        if let Some(target) = &query.address {
            target.validate()?;
        }
        let projection = query.projection.as_deref().unwrap_or("entries");
        if !matches!(projection, "entries" | "spills" | "status") {
            return Err(KernelError::new(
                "FORMULA_VALUE",
                "Unknown inspection projection",
            ));
        }
        let limit = query.limit.unwrap_or(512);
        if !(1..=4096).contains(&limit) {
            return Err(KernelError::new(
                "INSPECTION_QUERY_INVALID",
                "Inspection limit must be between 1 and 4096",
            ));
        }
        let session = evaluator::Session::new(self, reader, services);
        let bounds = inspection_bounds(query)?;
        let mut entries = Vec::new();
        let mut next_cursor = None;
        if projection == "entries" {
            let selected: Vec<_> = self
                .formulas
                .range(bounds.clone())
                .filter(|(cell, _)| query.address.as_ref().is_none_or(|target| target == *cell))
                .filter(|(cell, _)| {
                    query
                        .sheet_id
                        .as_ref()
                        .is_none_or(|sheet| sheet == &cell.sheet_id)
                })
                .take(limit + 1)
                .collect();
            if selected.len() > limit {
                next_cursor = selected
                    .get(limit - 1)
                    .map(|(cell, _)| serde_json::to_string(*cell).expect("cell addresses serialize"));
            }
            for (cell, formula) in selected.into_iter().take(limit) {
                let value = session.cell_output(cell)?;
                let mut dependencies = dependency_values(&formula.ast);
                if let Some(dynamic) = session.dynamic.borrow().get(cell) {
                    append_range_dependencies(&mut dependencies, dynamic);
                }
                entries.push(FormulaCellEntry {
                    address: cell.clone(),
                    formula: formula.source.clone(),
                    value,
                    dependencies,
                });
            }
        }
        session.validate_revision()?;
        let dependents = query
            .address
            .as_ref()
            .map_or_else(Vec::new, |target| dependents(self, target));
        let spills = if projection == "spills" {
            let selected: Vec<_> = self
                .spills
                .range(bounds)
                .map(|(cell, _)| cell)
                .filter(|cell| {
                    query
                        .sheet_id
                        .as_ref()
                        .is_none_or(|sheet| sheet == &cell.sheet_id)
                })
                .take(limit + 1)
                .collect();
            let mut values = Vec::new();
            for cell in selected.iter().take(limit) {
                let value = session.cell_output(cell)?;
                if let Some(spill) = spill_projection(cell, &value) {
                    values.push(spill);
                }
            }
            if selected.len() > limit {
                next_cursor = selected
                    .get(limit - 1)
                    .map(|cell| serde_json::to_string(*cell).expect("cell addresses serialize"));
            }
            values
        } else {
            Vec::new()
        };
        Ok(FormulaInspection {
            revision: reader.revision(),
            generation: self.generation(),
            entries,
            dependents,
            spills,
            pending_recalculation: !self.dirty.is_empty(),
            next_cursor,
        })
    }

    pub fn trace(
        &self,
        address: &CellAddress,
        reader: &dyn CellReader,
    ) -> KernelResult<FormulaTrace> {
        self.trace_with_services(address, reader, &DEFAULT_SERVICES)
    }
    pub fn trace_with_services(
        &self,
        address: &CellAddress,
        reader: &dyn CellReader,
        services: &dyn CalculationServices,
    ) -> KernelResult<FormulaTrace> {
        address.validate()?;
        let entry = self.formulas.get(address).ok_or_else(|| {
            KernelError::new("#VALUE!", "Formula trace requires an authored formula")
        })?;
        let session = evaluator::Session::new_traced(self, reader, services);
        session.force(address);
        let value = session.cell_output(address)?;
        let mut steps = session
            .trace_steps
            .into_inner()
            .into_iter()
            .map(|step| FormulaEvaluationTraceStep {
                expression: step.expression,
                value: step.value,
            })
            .collect::<Vec<_>>();
        if steps.is_empty() {
            steps.push(FormulaEvaluationTraceStep {
                expression: entry.source.clone(),
                value: value.clone(),
            });
        }
        Ok(FormulaTrace { value, steps })
    }

    pub fn spill_value(
        &self,
        address: &CellAddress,
        reader: &dyn CellReader,
    ) -> KernelResult<SpillValue> {
        self.spill_value_with_services(address, reader, &DEFAULT_SERVICES)
    }
    pub fn spill_value_with_services(
        &self,
        address: &CellAddress,
        reader: &dyn CellReader,
        services: &dyn CalculationServices,
    ) -> KernelResult<SpillValue> {
        address.validate()?;
        let value = evaluator::Session::new(self, reader, services).spill_scalar(address)?;
        Ok(SpillValue {
            is_spill: value.is_some(),
            value: value.map(FormulaValue::Scalar),
        })
    }
}

impl FormulaRuntime {
    pub fn evaluate_with_overrides(
        &self,
        formula: &str,
        current: &CellAddress,
        reader: &dyn CellReader,
        overrides: &BTreeMap<CellAddress, Scalar>,
    ) -> KernelResult<FormulaValue> {
        self.evaluate_with_overrides_and_services(
            formula,
            current,
            reader,
            overrides,
            &DEFAULT_SERVICES,
        )
    }
    pub fn evaluate_with_overrides_and_services(
        &self,
        formula: &str,
        current: &CellAddress,
        reader: &dyn CellReader,
        overrides: &BTreeMap<CellAddress, Scalar>,
        services: &dyn CalculationServices,
    ) -> KernelResult<FormulaValue> {
        let session = evaluator::Session::with_overrides(self, reader, services, overrides)?;
        let mut ast = parser::parse(formula, current)?;
        self.resolve_sheets(&mut ast)?;
        session.output(&ast, current)
    }
}

pub fn category(name: &str) -> &'static str {
    let name = name.to_ascii_uppercase();
    if matches!(
        name.as_str(),
        "IF" | "IFS"
            | "IFERROR"
            | "IFNA"
            | "SWITCH"
            | "AND"
            | "OR"
            | "LET"
            | "LAMBDA"
            | "MAP"
            | "REDUCE"
            | "SCAN"
            | "BYROW"
            | "BYCOL"
            | "MAKEARRAY"
            | "ISOMITTED"
    ) {
        "logical"
    } else if matches!(
        name.as_str(),
        "CONCAT"
            | "CONCATENATE"
            | "TEXTJOIN"
            | "LEN"
            | "LEFT"
            | "RIGHT"
            | "MID"
            | "LOWER"
            | "UPPER"
            | "TRIM"
            | "TEXT"
            | "VALUE"
            | "EXACT"
    ) {
        "text"
    } else if date_stat_functions::FUNCTIONS.contains(&name.as_str()) {
        "date-time"
    } else if scalar_functions::FUNCTIONS.contains(&name.as_str()) {
        "math-trig"
    } else if array_functions::FUNCTIONS.contains(&name.as_str()) {
        "lookup-reference"
    } else if evaluator::FUNCTIONS.contains(&name.as_str()) {
        "more-functions"
    } else {
        "unknown"
    }
}
