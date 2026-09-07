//! Reference-aware calculation. Host failures abort; Excel errors are cell values.
use kernel_core::{
    CellAddress, CellReader, KernelError, KernelResult, MAX_COLUMNS, MAX_ROWS, RangeRef, Scalar,
};
use serde::{Deserialize, Serialize};
use std::{
    cell::{Cell, RefCell},
    collections::{BTreeMap, BTreeSet},
    rc::Rc,
};
mod array_functions;
mod date_stat_functions;
mod dependency;
pub mod editor;
mod evaluator;
mod inspection;
pub mod parser;
pub mod references;
mod scalar_functions;
use dependency::DependencyIndex;
pub use inspection::*;
use parser::Expr;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum FormulaValue {
    Scalar(Scalar),
    Array(Vec<Vec<Scalar>>),
}
impl FormulaValue {
    pub fn scalar(&self) -> Scalar {
        match self {
            Self::Scalar(v) => v.clone(),
            Self::Array(v) => v
                .first()
                .and_then(|r| r.first())
                .cloned()
                .unwrap_or_default(),
        }
    }
    pub fn matrix(&self) -> Vec<Vec<Scalar>> {
        match self {
            Self::Scalar(v) => vec![vec![v.clone()]],
            Self::Array(v) => v.clone(),
        }
    }
    pub fn error(code: &str, message: impl Into<String>) -> Self {
        Self::Scalar(Scalar::error(code, message))
    }
    pub fn shape(&self) -> (usize, usize) {
        match self {
            Self::Scalar(_) => (1, 1),
            Self::Array(v) => (v.len(), v.first().map_or(0, Vec::len)),
        }
    }
    pub fn is_error(&self) -> bool {
        matches!(self, Self::Scalar(Scalar::Error(_)))
    }
}
pub(crate) fn num(v: &Scalar) -> KernelResult<f64> {
    match v {
        Scalar::Number(n) => Ok(*n),
        Scalar::Boolean(v) => Ok(if *v { 1. } else { 0. }),
        Scalar::Null => Ok(0.),
        Scalar::Text(v) => v
            .trim()
            .parse::<f64>()
            .ok()
            .filter(|n| n.is_finite())
            .ok_or_else(|| KernelError::new("#VALUE!", "Expected a number")),
        Scalar::Error(e) => Err(KernelError::new(&e.code, &e.message)),
    }
}
pub(crate) fn truth(v: &Scalar) -> KernelResult<bool> {
    match v {
        Scalar::Boolean(v) => Ok(*v),
        Scalar::Number(n) => Ok(*n != 0.),
        Scalar::Null => Ok(false),
        Scalar::Text(v) if v.eq_ignore_ascii_case("TRUE") => Ok(true),
        Scalar::Text(v) if v.eq_ignore_ascii_case("FALSE") => Ok(false),
        Scalar::Text(_) => Err(KernelError::new("#VALUE!", "Text is not a logical value")),
        Scalar::Error(e) => Err(KernelError::new(&e.code, &e.message)),
    }
}
pub(crate) fn text(v: &Scalar) -> String {
    match v {
        Scalar::Null => String::new(),
        Scalar::Boolean(true) => "TRUE".into(),
        Scalar::Boolean(false) => "FALSE".into(),
        Scalar::Number(v) => v.to_string(),
        Scalar::Text(v) => v.clone(),
        Scalar::Error(e) => e.code.clone(),
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CalculationContext {
    pub date1904: bool,
    /// Transaction instant expressed as an Excel serial, identical on both hosts.
    pub now_serial: Option<f64>,
    pub random_seed: u64,
    pub max_steps: u64,
    pub max_array_cells: u64,
    pub max_recursion: u32,
}
impl Default for CalculationContext {
    fn default() -> Self {
        Self {
            date1904: false,
            now_serial: None,
            random_seed: 1,
            max_steps: 100_000_000,
            max_array_cells: 4_000_000,
            max_recursion: 256,
        }
    }
}
#[derive(Debug, Clone, Copy, Default)]
pub struct FormulaVisibility {
    pub filter_hidden: bool,
    pub manual_hidden: bool,
    pub outline_hidden: bool,
}
pub trait CalculationServices {
    fn checkpoint(&self, _steps: u64) -> KernelResult<()> {
        Ok(())
    }
    fn visibility(&self, address: &CellAddress) -> KernelResult<FormulaVisibility> {
        Err(KernelError::new(
            "FORMULA_VISIBILITY_UNAVAILABLE",
            "Canonical visibility is required for SUBTOTAL/AGGREGATE",
        )
        .at(format!(
            "{}:{}:{}",
            address.sheet_id, address.row, address.column
        )))
    }
    fn external(&self, name: &str, _args: &[FormulaValue]) -> KernelResult<FormulaValue> {
        Err(KernelError::new(
            "UNSUPPORTED_FEATURE",
            format!("External formula capability {name} is unavailable"),
        )
        .at(name))
    }
}
struct DefaultServices;
impl CalculationServices for DefaultServices {}
static DEFAULT_SERVICES: DefaultServices = DefaultServices;
#[derive(Clone)]
struct FormulaEntry {
    ast: Expr,
    source: String,
}
#[derive(Clone)]
pub struct FormulaRuntime {
    sheet_ids: BTreeMap<String, String>,
    generation: u64,
    formulas: BTreeMap<CellAddress, FormulaEntry>,
    /// Explicit overrides only; never populated by a page scan.
    values: BTreeMap<CellAddress, Scalar>,
    cached: BTreeMap<CellAddress, FormulaValue>,
    index: DependencyIndex,
    dirty: BTreeSet<CellAddress>,
    volatile: BTreeSet<CellAddress>,
    names: BTreeMap<String, Expr>,
    tables: BTreeMap<String, RangeRef>,
    spills: BTreeMap<CellAddress, RangeRef>,
    spill_index: DependencyIndex,
    pub context: CalculationContext,
}
impl FormulaRuntime {
    pub fn new(default_sheet_id: impl Into<String>) -> Self {
        let id = default_sheet_id.into();
        Self {
            sheet_ids: BTreeMap::from([(id.to_uppercase(), id)]),
            generation: 0,
            formulas: BTreeMap::new(),
            values: BTreeMap::new(),
            cached: BTreeMap::new(),
            index: DependencyIndex::default(),
            dirty: BTreeSet::new(),
            volatile: BTreeSet::new(),
            names: BTreeMap::new(),
            tables: BTreeMap::new(),
            spills: BTreeMap::new(),
            spill_index: DependencyIndex::default(),
            context: CalculationContext::default(),
        }
    }
    pub fn set_value(&mut self, address: CellAddress, value: Scalar) -> KernelResult<()> {
        address.validate()?;
        value.validate()?;
        self.remove_formula(&address);
        self.values.insert(address.clone(), value);
        self.invalidate(&address);
        Ok(())
    }
    pub fn set_formula(&mut self, address: CellAddress, formula: &str) -> KernelResult<()> {
        address.validate()?;
        let mut ast = parser::parse(formula, &address)?;
        self.resolve_sheets(&mut ast)?;
        self.index.replace(&address, &ast);
        self.values.remove(&address);
        if dependency::is_volatile(&ast) {
            self.volatile.insert(address.clone());
        } else {
            self.volatile.remove(&address);
        }
        self.formulas.insert(
            address.clone(),
            FormulaEntry {
                ast,
                source: formula.into(),
            },
        );
        self.dirty.insert(address.clone());
        self.invalidate(&address);
        Ok(())
    }
    pub fn remove_formula(&mut self, address: &CellAddress) {
        if let Some(range) = self.spills.remove(address) {
            for dependent in self.index.affected_range(&range) {
                self.dirty.insert(dependent.clone());
                self.invalidate(&dependent);
            }
        }
        self.spill_index.remove(address);
        self.index.remove(address);
        self.formulas.remove(address);
        self.cached.remove(address);
        self.volatile.remove(address);
        self.dirty.remove(address);
    }
    pub fn invalidate(&mut self, address: &CellAddress) {
        let mut queue = vec![address.clone()];
        let mut seen = BTreeSet::new();
        while let Some(changed) = queue.pop() {
            if !seen.insert(changed.clone()) {
                continue;
            }
            for dependent in self.index.affected(&changed) {
                if self.dirty.insert(dependent.clone()) {
                    queue.push(dependent);
                }
            }
        }
    }
    pub fn invalidate_range(&mut self, range: &RangeRef) -> KernelResult<()> {
        range.validate()?;
        for address in self.index.affected_range(range) {
            self.dirty.insert(address.clone());
            self.invalidate(&address);
        }
        Ok(())
    }
    pub fn define_name(
        &mut self,
        name: &str,
        formula: &str,
        scope: &CellAddress,
    ) -> KernelResult<()> {
        let key = name.to_uppercase();
        let mut ast = parser::parse(formula, scope)?;
        self.resolve_sheets(&mut ast)?;
        self.names.insert(key.clone(), ast);
        for a in self.index.named(&key) {
            self.dirty.insert(a.clone());
            self.invalidate(&a);
        }
        Ok(())
    }
    pub fn define_table_reference(&mut self, name: &str, range: RangeRef) -> KernelResult<()> {
        range.validate()?;
        let key = name.to_uppercase();
        self.tables.insert(key.clone(), range);
        for a in self.index.named(&key) {
            self.dirty.insert(a.clone());
            self.invalidate(&a);
        }
        Ok(())
    }
    pub fn evaluate(
        &self,
        formula: &str,
        current: &CellAddress,
        reader: &dyn CellReader,
    ) -> KernelResult<FormulaValue> {
        self.evaluate_with_services(formula, current, reader, &DEFAULT_SERVICES)
    }
    pub fn register_sheet(&mut self, name: &str, id: &str) -> KernelResult<()> {
        if name.is_empty() || id.is_empty() {
            return Err(KernelError::new(
                "SHEET_IDENTITY_INVALID",
                "Sheet name and id are required",
            ));
        }
        self.sheet_ids.insert(name.to_uppercase(), id.into());
        self.sheet_ids.insert(id.to_uppercase(), id.into());
        Ok(())
    }
    fn resolved_range(&self, range: &RangeRef) -> KernelResult<RangeRef> {
        let mut range = range.clone();
        if let Some(id) = self.sheet_ids.get(&range.sheet_id.to_uppercase()) {
            range.sheet_id = id.clone();
        } else if !self.sheet_ids.is_empty() {
            return Err(KernelError::new(
                "#REF!",
                format!("Unknown worksheet {}", range.sheet_id),
            ));
        }
        Ok(range)
    }
    fn resolve_sheets(&self, expr: &mut Expr) -> KernelResult<()> {
        match expr {
            Expr::Reference(range) => *range = self.resolved_range(range)?,
            Expr::Array(rows) => {
                for value in rows.iter_mut().flatten() {
                    self.resolve_sheets(value)?
                }
            }
            Expr::Unary(_, value) | Expr::Spill(value) => self.resolve_sheets(value)?,
            Expr::Binary(_, left, right) => {
                self.resolve_sheets(left)?;
                self.resolve_sheets(right)?
            }
            Expr::Call(_, args) => {
                for arg in args {
                    self.resolve_sheets(arg)?
                }
            }
            Expr::Invoke(callee, args) => {
                self.resolve_sheets(callee)?;
                for arg in args {
                    self.resolve_sheets(arg)?
                }
            }
            _ => {}
        }
        Ok(())
    }
    pub fn generation(&self) -> u64 {
        self.generation
    }
    pub fn set_context(&mut self, context: CalculationContext) -> KernelResult<()> {
        if context.max_steps == 0
            || context.max_array_cells == 0
            || context.max_recursion == 0
            || context.now_serial.is_some_and(|n| !n.is_finite())
        {
            return Err(KernelError::new(
                "CALCULATION_CONTEXT_INVALID",
                "Calculation context has invalid bounds",
            ));
        }
        if context.date1904 != self.context.date1904 {
            self.dirty.extend(self.formulas.keys().cloned());
        }
        self.context = context;
        Ok(())
    }
    pub fn evaluate_with_services(
        &self,
        formula: &str,
        current: &CellAddress,
        reader: &dyn CellReader,
        services: &dyn CalculationServices,
    ) -> KernelResult<FormulaValue> {
        current.validate()?;
        let mut ast = parser::parse(formula, current)?;
        self.resolve_sheets(&mut ast)?;
        evaluator::Session::new(self, reader, services).output(&ast, current)
    }
    pub fn recalculate(
        &mut self,
        reader: &dyn CellReader,
    ) -> KernelResult<BTreeMap<CellAddress, FormulaValue>> {
        self.recalculate_with_services(reader, &DEFAULT_SERVICES)
    }
    pub fn recalculate_with_services(
        &mut self,
        reader: &dyn CellReader,
        services: &dyn CalculationServices,
    ) -> KernelResult<BTreeMap<CellAddress, FormulaValue>> {
        for a in self.volatile.clone() {
            self.dirty.insert(a.clone());
            self.invalidate(&a);
        }
        let session = evaluator::Session::new(self, reader, services);
        let mut result = BTreeMap::new();
        for address in &self.dirty {
            if self.formulas.contains_key(address) {
                result.insert(address.clone(), session.cell_output(address)?);
            }
        }
        // Spill footprints are dependencies too. New/resized/removed footprints
        // schedule only formulas whose indexed ranges intersect those footprints.
        let mut rounds = 0;
        loop {
            let roots = std::mem::take(&mut *session.spill_changed.borrow_mut());
            if roots.is_empty() {
                break;
            }
            rounds += 1;
            if rounds > self.context.max_recursion {
                return Err(KernelError::new(
                    "RESOURCE_BUDGET_EXCEEDED",
                    "Spill dependency graph did not converge",
                ));
            }
            let mut affected = BTreeSet::new();
            for root in roots {
                if let Some(old) = self.spills.get(&root) {
                    affected.extend(self.index.affected_range(old));
                }
                if let Some(Some(new)) = session.spill_updates.borrow().get(&root) {
                    affected.extend(self.index.affected_range(new));
                }
                affected.remove(&root);
            }
            let mut queue: Vec<_> = affected.iter().cloned().collect();
            while let Some(a) = queue.pop() {
                for dependent in self.index.affected(&a) {
                    if affected.insert(dependent.clone()) {
                        queue.push(dependent);
                    }
                }
            }
            for a in &affected {
                session.force(a);
            }
            for a in affected {
                if self.formulas.contains_key(&a) {
                    result.insert(a.clone(), session.cell_output(&a)?);
                }
            }
        }
        session.validate_revision()?;
        let next_generation = if result.is_empty() {
            self.generation
        } else {
            self.generation.checked_add(1).ok_or_else(|| {
                KernelError::new("GENERATION_EXHAUSTED", "Calculation generation overflow")
            })?
        };
        let dynamic = session.dynamic.into_inner();
        let spills = session.spill_updates.into_inner();
        for (a, ranges) in dynamic {
            if let Some(entry) = self.formulas.get(&a) {
                self.index.replace(&a, &entry.ast);
                self.index.add_dynamic(&a, ranges);
            }
        }
        for (a, range) in spills {
            self.spill_index.remove(&a);
            match range {
                Some(range) => {
                    self.spill_index.add_dynamic(&a, vec![range.clone()]);
                    self.spills.insert(a, range);
                }
                None => {
                    self.spills.remove(&a);
                }
            }
        }
        for (a, value) in &result {
            self.cached.insert(a.clone(), value.clone());
        }
        self.dirty.clear();
        self.generation = next_generation;
        Ok(result)
    }
    pub fn dependencies(&self, address: &CellAddress) -> Option<&BTreeSet<CellAddress>> {
        self.index.points(address)
    }
    pub fn formula_count(&self) -> usize {
        self.formulas.len()
    }
    pub fn dirty_count(&self) -> usize {
        self.dirty.len()
    }
}
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FunctionCapability {
    pub id: String,
    pub cost: &'static str,
    pub streaming: bool,
    pub volatile: bool,
    pub status: &'static str,
    pub category: &'static str,
}
pub fn function_capabilities() -> Vec<FunctionCapability> {
    let mut functions = BTreeMap::new();
    for (names, _) in [
        (scalar_functions::FUNCTIONS, "math-trig"),
        (date_stat_functions::FUNCTIONS, "date-time"),
        (array_functions::FUNCTIONS, "lookup-reference"),
        (evaluator::FUNCTIONS, "more-functions"),
    ] {
        for name in names {
            let volatile = matches!(
                *name,
                "NOW" | "TODAY" | "RAND" | "RANDBETWEEN" | "RANDARRAY" | "OFFSET" | "INDIRECT"
            );
            let streaming = evaluator::STREAMING.contains(name);
            functions.insert(
                *name,
                FunctionCapability {
                    id: (*name).into(),
                    cost: if volatile {
                        "volatile"
                    } else if streaming {
                        "range"
                    } else {
                        "scalar"
                    },
                    streaming,
                    volatile,
                    status: "native",
                    category: inspection::category(name),
                },
            );
        }
    }
    functions.into_values().collect()
}
