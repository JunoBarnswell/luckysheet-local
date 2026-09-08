use kernel_core::{Cell, CellAddress, CellReader, KernelError, KernelResult, RangeRef, Scalar};
use kernel_formula::{DefinedNameScope, FormulaRuntime, FormulaTable, parser::Expr};
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, BTreeSet};

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FieldDescriptor {
    pub field_id: String,
    pub name: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CalculatedField {
    pub field_id: String,
    pub name: String,
    pub formula: String,
}

/// A calculated item is retained as a source-row operation. It cannot be
/// evaluated from aggregate states because its references select members and
/// then aggregate source rows for each Values field.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CalculatedItem {
    pub field_id: String,
    pub target_field_id: String,
    pub name: String,
    pub formula: String,
}

#[derive(Debug)]
pub struct Evaluator {
    fields: Vec<FieldDescriptor>,
    definitions: BTreeMap<String, CalculatedField>,
    ordered: Vec<String>,
}

impl Evaluator {
    pub fn new(
        fields: Vec<FieldDescriptor>,
        calculated_fields: Vec<CalculatedField>,
    ) -> KernelResult<Self> {
        let mut all = fields;
        let mut definitions = BTreeMap::new();
        for definition in calculated_fields {
            if definition.field_id.trim().is_empty()
                || definition.name.trim().is_empty()
                || definition.formula.trim().is_empty()
            {
                return Err(KernelError::new(
                    "PIVOT_CALCULATED_FIELD_INVALID",
                    "Calculated field identity, name, and formula are required",
                )
                .at(definition.field_id));
            }
            if definitions
                .insert(definition.field_id.clone(), definition.clone())
                .is_some()
            {
                return Err(KernelError::new(
                    "PIVOT_CALCULATED_FIELD_DUPLICATE",
                    "Calculated field identity is duplicated",
                )
                .at(definition.field_id));
            }
            if !all
                .iter()
                .any(|field| field.field_id == definition.field_id)
            {
                all.push(FieldDescriptor {
                    field_id: definition.field_id,
                    name: definition.name,
                });
            }
        }
        validate_aliases(&all)?;
        let mut dependencies = BTreeMap::<String, Vec<String>>::new();
        for definition in definitions.values() {
            dependencies.insert(
                definition.field_id.clone(),
                references(
                    &definition.formula,
                    &all,
                    &definitions,
                    &definition.field_id,
                )?,
            );
        }
        let mut state = BTreeMap::<String, bool>::new();
        let mut ordered = Vec::new();
        for id in definitions.keys() {
            visit(id, &dependencies, &mut state, &mut ordered, &mut Vec::new())?;
        }
        Ok(Self {
            fields: all,
            definitions,
            ordered,
        })
    }

    pub fn has(&self, field_id: &str) -> bool {
        self.definitions.contains_key(field_id)
    }

    /// Evaluate one calculated field against the already aggregated Values
    /// states. Missing source states are canonical Null; no text-to-number
    /// coercion is introduced here.
    pub fn evaluate(
        &self,
        field_id: &str,
        states: &BTreeMap<String, Scalar>,
    ) -> KernelResult<Scalar> {
        if !self.has(field_id) {
            return Err(KernelError::new(
                "PIVOT_CALCULATED_FIELD_UNKNOWN",
                "Calculated field is not defined",
            )
            .at(field_id));
        }
        let values = self.evaluate_all(states)?;
        values.get(field_id).cloned().ok_or_else(|| {
            KernelError::new(
                "PIVOT_CALCULATED_FIELD_UNKNOWN",
                "Calculated field evaluation produced no value",
            )
            .at(field_id)
        })
    }

    pub fn evaluate_all(
        &self,
        states: &BTreeMap<String, Scalar>,
    ) -> KernelResult<BTreeMap<String, Scalar>> {
        let mut runtime = FormulaRuntime::new("pivot-summary");
        let reader = StateReader { values: states };
        for (index, field) in self.fields.iter().enumerate() {
            let address = CellAddress {
                sheet_id: "pivot-summary".into(),
                row: 0,
                column: index as u32,
            };
            runtime.set_value(
                address.clone(),
                states.get(&field.field_id).cloned().unwrap_or(Scalar::Null),
            )?;
            let reference = format!("=${}1", column_name(index as u32));
            runtime.define_name(
                &field.field_id,
                &reference,
                DefinedNameScope::Workbook,
                &CellAddress {
                    sheet_id: "pivot-summary".into(),
                    row: 1,
                    column: 0,
                },
            )?;
            runtime.define_name(
                &field.name,
                &reference,
                DefinedNameScope::Workbook,
                &CellAddress {
                    sheet_id: "pivot-summary".into(),
                    row: 1,
                    column: 0,
                },
            )?;
            let range = RangeRef {
                sheet_id: "pivot-summary".into(),
                start_row: 0,
                end_row: 0,
                start_column: index as u32,
                end_column: index as u32,
            };
            for name in [&field.field_id, &field.name] {
                runtime.define_table(FormulaTable {
                    name: name.clone(),
                    range: range.clone(),
                    has_header_row: false,
                    has_total_row: false,
                    columns: vec![name.clone()],
                })?;
            }
        }
        let mut output = BTreeMap::new();
        for id in &self.ordered {
            let definition = self.definitions.get(id).expect("ordered definition");
            let address = CellAddress {
                sheet_id: "pivot-summary".into(),
                row: 1,
                column: 0,
            };
            let value = runtime
                .evaluate(&definition.formula, &address, &reader)?
                .scalar();
            if let Scalar::Error(error) = &value {
                if error.code == "#NAME?" || error.code == "#REF!" {
                    return Err(KernelError::new(
                        "PIVOT_CALCULATED_FIELD_REFERENCE",
                        "Calculated field contains an invalid reference",
                    )
                    .at(&definition.field_id));
                }
            }
            let index = self
                .fields
                .iter()
                .position(|field| field.field_id == *id)
                .ok_or_else(|| {
                    KernelError::new(
                        "PIVOT_CALCULATED_FIELD_INVALID",
                        "Calculated field descriptor is missing",
                    )
                    .at(id)
                })?;
            runtime.set_value(
                CellAddress {
                    sheet_id: "pivot-summary".into(),
                    row: 0,
                    column: index as u32,
                },
                value.clone(),
            )?;
            output.insert(id.clone(), value);
        }
        Ok(output)
    }
}

struct StateReader<'a> {
    values: &'a BTreeMap<String, Scalar>,
}
impl CellReader for StateReader<'_> {
    fn revision(&self) -> u64 {
        0
    }
    fn read_cell(&self, address: &CellAddress) -> KernelResult<Option<Cell>> {
        if address.sheet_id != "pivot-summary" {
            return Err(KernelError::new(
                "PIVOT_CALCULATED_FIELD_REFERENCE",
                "Calculated field references another sheet",
            ));
        }
        Ok(Some(Cell {
            value: self
                .values
                .get(&address.column.to_string())
                .cloned()
                .unwrap_or(Scalar::Null),
            formula: None,
            metadata: BTreeMap::new(),
        }))
    }
    fn read_range(
        &self,
        range: &RangeRef,
        visitor: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        for row in range.start_row..=range.end_row {
            for column in range.start_column..=range.end_column {
                visitor(
                    CellAddress {
                        sheet_id: range.sheet_id.clone(),
                        row,
                        column,
                    },
                    Cell {
                        value: Scalar::Null,
                        formula: None,
                        metadata: BTreeMap::new(),
                    },
                )?;
            }
        }
        Ok(())
    }
}

fn validate_aliases(fields: &[FieldDescriptor]) -> KernelResult<()> {
    let mut aliases = BTreeMap::<String, String>::new();
    for field in fields {
        for alias in [&field.field_id, &field.name] {
            let key = alias.to_ascii_uppercase();
            if let Some(previous) = aliases.insert(key.clone(), field.field_id.clone()) {
                if previous != field.field_id {
                    return Err(KernelError::new(
                        "PIVOT_CALCULATED_FIELD_AMBIGUOUS",
                        "Field id or name is ambiguous",
                    )
                    .at(alias));
                }
            }
        }
    }
    Ok(())
}

fn column_name(mut column: u32) -> String {
    let mut result = String::new();
    loop {
        result.insert(0, (b'A' + (column % 26) as u8) as char);
        if column < 26 {
            return result;
        }
        column = column / 26 - 1;
    }
}

fn references(
    formula: &str,
    fields: &[FieldDescriptor],
    definitions: &BTreeMap<String, CalculatedField>,
    owner: &str,
) -> KernelResult<Vec<String>> {
    let address = CellAddress {
        sheet_id: "pivot-summary".into(),
        row: 1,
        column: 0,
    };
    let ast = kernel_formula::parser::parse(formula, &address).map_err(|e| {
        KernelError::new("PIVOT_CALCULATED_FIELD_FORMULA_INVALID", e.message).at(owner)
    })?;
    let mut aliases = BTreeMap::<String, String>::new();
    for field in fields {
        aliases.insert(field.field_id.to_ascii_uppercase(), field.field_id.clone());
        aliases.insert(field.name.to_ascii_uppercase(), field.field_id.clone());
    }
    let mut found = BTreeSet::new();
    fn walk(
        expr: &Expr,
        aliases: &BTreeMap<String, String>,
        defs: &BTreeMap<String, CalculatedField>,
        found: &mut BTreeSet<String>,
        owner: &str,
    ) -> KernelResult<()> {
        match expr {
            Expr::Name(name) => {
                let id = aliases.get(&name.to_ascii_uppercase()).ok_or_else(|| {
                    KernelError::new(
                        "PIVOT_CALCULATED_FIELD_REFERENCE",
                        format!("Unknown calculated field reference: {name}"),
                    )
                    .at(owner)
                })?;
                if defs.contains_key(id) {
                    found.insert(id.clone());
                }
            }
            Expr::Structured(reference) => {
                let id = aliases
                    .get(&reference.table_name.to_ascii_uppercase())
                    .ok_or_else(|| {
                        KernelError::new(
                            "PIVOT_CALCULATED_FIELD_REFERENCE",
                            format!(
                                "Unknown calculated field reference: {}",
                                reference.table_name
                            ),
                        )
                        .at(owner)
                    })?;
                if defs.contains_key(id) {
                    found.insert(id.clone());
                }
            }
            Expr::Unary(_, e) | Expr::Spill(e) => walk(e, aliases, defs, found, owner)?,
            Expr::Binary(_, a, b) => {
                walk(a, aliases, defs, found, owner)?;
                walk(b, aliases, defs, found, owner)?;
            }
            Expr::Array(rows) => {
                for row in rows {
                    for e in row {
                        walk(e, aliases, defs, found, owner)?;
                    }
                }
            }
            Expr::Call(_, args) | Expr::Invoke(_, args) => {
                for e in args {
                    walk(e, aliases, defs, found, owner)?;
                }
            }
            Expr::Scalar(_) | Expr::Reference(_) | Expr::Missing => {}
        }
        Ok(())
    }
    walk(&ast, &aliases, definitions, &mut found, owner)?;
    Ok(found.into_iter().collect())
}

fn visit(
    id: &str,
    dependencies: &BTreeMap<String, Vec<String>>,
    state: &mut BTreeMap<String, bool>,
    ordered: &mut Vec<String>,
    path: &mut Vec<String>,
) -> KernelResult<()> {
    if state.get(id) == Some(&true) {
        return Ok(());
    }
    if state.get(id) == Some(&false) {
        path.push(id.into());
        return Err(KernelError::new(
            "PIVOT_CALCULATED_FIELD_CYCLE",
            format!("Calculated field dependency cycle: {}", path.join(" -> ")),
        )
        .at(id));
    }
    state.insert(id.into(), false);
    path.push(id.into());
    for dependency in dependencies.get(id).into_iter().flatten() {
        visit(dependency, dependencies, state, ordered, path)?;
    }
    path.pop();
    state.insert(id.into(), true);
    ordered.push(id.into());
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn evaluates_dependencies_from_states() {
        let evaluator = Evaluator::new(
            vec![FieldDescriptor {
                field_id: "sales".into(),
                name: "Sales".into(),
            }],
            vec![CalculatedField {
                field_id: "gross".into(),
                name: "Gross".into(),
                formula: "=Sales*2".into(),
            }],
        )
        .unwrap();
        let mut states = BTreeMap::new();
        states.insert("sales".into(), Scalar::Number(4.0));
        assert_eq!(
            evaluator.evaluate("gross", &states).unwrap(),
            Scalar::Number(8.0)
        );
    }
    #[test]
    fn rejects_unknown_reference_and_cycle() {
        let unknown = Evaluator::new(
            vec![FieldDescriptor {
                field_id: "sales".into(),
                name: "Sales".into(),
            }],
            vec![CalculatedField {
                field_id: "gross".into(),
                name: "Gross".into(),
                formula: "=Missing+1".into(),
            }],
        );
        assert_eq!(
            unknown.unwrap_err().code,
            "PIVOT_CALCULATED_FIELD_REFERENCE"
        );
        let cycle = Evaluator::new(
            vec![],
            vec![
                CalculatedField {
                    field_id: "a".into(),
                    name: "A".into(),
                    formula: "=B".into(),
                },
                CalculatedField {
                    field_id: "b".into(),
                    name: "B".into(),
                    formula: "=A".into(),
                },
            ],
        );
        assert_eq!(cycle.unwrap_err().code, "PIVOT_CALCULATED_FIELD_CYCLE");
    }
}
