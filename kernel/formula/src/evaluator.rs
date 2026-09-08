use super::*;
type Env = BTreeMap<String, Value>;
#[derive(Clone)]
enum Value {
    Data(FormulaValue),
    Reference(RangeRef),
    Lambda(Rc<Lambda>),
    Missing,
    Union(Vec<Value>),
}
#[derive(Clone)]
struct Lambda {
    params: Vec<String>,
    body: Expr,
    env: Env,
    recursive_name: Option<String>,
}
impl Value {
    fn scalar(v: Scalar) -> Self {
        Self::Data(FormulaValue::Scalar(v))
    }
    fn error(code: &str, msg: impl Into<String>) -> Self {
        Self::Data(FormulaValue::error(code, msg))
    }
}
pub const STREAMING: &[&str] = &[
    "SUM",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "AVERAGE",
    "MIN",
    "MAX",
    "PRODUCT",
    "VAR",
    "VARP",
    "STDEV",
    "STDEVP",
    "SUMIF",
    "SUMIFS",
    "COUNTIF",
    "COUNTIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
    "MAXIFS",
    "MINIFS",
    "SUBTOTAL",
    "AGGREGATE",
    "SUMPRODUCT",
    "AND",
    "OR",
];
pub const FUNCTIONS: &[&str] = &[
    "SUM",
    "COUNT",
    "COUNTA",
    "COUNTBLANK",
    "AVERAGE",
    "MIN",
    "MAX",
    "PRODUCT",
    "VAR",
    "VARP",
    "STDEV",
    "STDEVP",
    "SUMIF",
    "SUMIFS",
    "COUNTIF",
    "COUNTIFS",
    "AVERAGEIF",
    "AVERAGEIFS",
    "MAXIFS",
    "MINIFS",
    "SUBTOTAL",
    "AGGREGATE",
    "SUMPRODUCT",
    "AND",
    "OR",
    "IF",
    "IFS",
    "IFERROR",
    "IFNA",
    "SWITCH",
    "CHOOSE",
    "LET",
    "LAMBDA",
    "MAP",
    "REDUCE",
    "SCAN",
    "BYROW",
    "BYCOL",
    "MAKEARRAY",
    "ROW",
    "COLUMN",
    "ADDRESS",
    "OFFSET",
    "INDIRECT",
    "NOW",
    "TODAY",
    "RAND",
    "RANDBETWEEN",
    "RANDARRAY",
    "ISOMITTED",
    "SJS.TABLE",
];
pub(crate) struct Session<'a> {
    runtime: &'a FormulaRuntime,
    reader: &'a dyn CellReader,
    services: &'a dyn CalculationServices,
    revision: u64,
    steps: Cell<u64>,
    depth: Cell<u32>,
    random_counts: RefCell<BTreeMap<CellAddress, u64>>,
    overrides: Option<&'a BTreeMap<CellAddress, Scalar>>,
    stack: RefCell<BTreeSet<CellAddress>>,
    memo: RefCell<BTreeMap<CellAddress, FormulaValue>>,
    pub dynamic: RefCell<BTreeMap<CellAddress, Vec<RangeRef>>>,
    pub spill_updates: RefCell<BTreeMap<CellAddress, Option<RangeRef>>>,
    pub spill_changed: RefCell<BTreeSet<CellAddress>>,
    spill_index: RefCell<DependencyIndex>,
    forced: RefCell<BTreeSet<CellAddress>>,
    tracing: bool,
    pub trace_steps: RefCell<Vec<crate::FormulaEvaluationTraceStep>>,
}
impl<'a> Session<'a> {
    pub fn new(
        runtime: &'a FormulaRuntime,
        reader: &'a dyn CellReader,
        services: &'a dyn CalculationServices,
    ) -> Self {
        Self {
            runtime,
            reader,
            services,
            revision: reader.revision(),
            steps: Cell::new(0),
            depth: Cell::new(0),
            random_counts: RefCell::new(BTreeMap::new()),
            overrides: None,
            stack: RefCell::new(BTreeSet::new()),
            memo: RefCell::new(BTreeMap::new()),
            dynamic: RefCell::new(BTreeMap::new()),
            spill_updates: RefCell::new(BTreeMap::new()),
            spill_changed: RefCell::new(BTreeSet::new()),
            spill_index: RefCell::new(DependencyIndex::default()),
            forced: RefCell::new(BTreeSet::new()),
            tracing: false,
            trace_steps: RefCell::new(Vec::new()),
        }
    }
    pub fn new_traced(
        runtime: &'a FormulaRuntime,
        reader: &'a dyn CellReader,
        services: &'a dyn CalculationServices,
    ) -> Self {
        let mut session = Self::new(runtime, reader, services);
        session.tracing = true;
        session
    }
    pub fn force(&self, address: &CellAddress) {
        self.memo.borrow_mut().remove(address);
        self.random_counts.borrow_mut().remove(address);
        self.forced.borrow_mut().insert(address.clone());
    }
    pub fn with_overrides(
        runtime: &'a FormulaRuntime,
        reader: &'a dyn CellReader,
        services: &'a dyn CalculationServices,
        overrides: &'a BTreeMap<CellAddress, Scalar>,
    ) -> KernelResult<Self> {
        let mut session = Self::new(runtime, reader, services);
        session.overrides = Some(overrides);
        let mut queue = Vec::new();
        for (address, value) in overrides {
            address.validate()?;
            value.validate()?;
            session.force(address);
            queue.push(address.clone());
        }
        while let Some(address) = queue.pop() {
            for dependent in runtime.index.affected(&address) {
                if !session.forced.borrow().contains(&dependent) {
                    session.force(&dependent);
                    queue.push(dependent);
                }
            }
        }
        Ok(session)
    }
    fn tick(&self) -> KernelResult<()> {
        let step = self.steps.get() + 1;
        self.steps.set(step);
        if step > self.runtime.context.max_steps {
            return Err(KernelError::new(
                "RESOURCE_BUDGET_EXCEEDED",
                "Formula operation budget exhausted",
            )
            .recover("increase-budget-or-reduce-calculation"));
        }
        if step == 1 || step % 256 == 0 {
            self.services.checkpoint(step)?;
            self.validate_revision()?;
        }
        Ok(())
    }
    pub fn validate_revision(&self) -> KernelResult<()> {
        if self.reader.revision() != self.revision {
            Err(
                KernelError::new("STALE_REVISION", "Formula input changed during calculation")
                    .recover("retry-current-revision"),
            )
        } else {
            Ok(())
        }
    }
    fn allocate(&self, rows: usize, columns: usize) -> KernelResult<()> {
        if rows
            .checked_mul(columns)
            .is_none_or(|n| n as u64 > self.runtime.context.max_array_cells)
        {
            Err(KernelError::new(
                "RESOURCE_BUDGET_EXCEEDED",
                "Formula array exceeds the transaction allocation budget",
            ))
        } else {
            Ok(())
        }
    }
    pub fn output(&self, expr: &Expr, current: &CellAddress) -> KernelResult<FormulaValue> {
        let v = self.eval(expr, current, &Env::new())?;
        let mut result = self.materialize(&v)?;
        match &mut result {
            FormulaValue::Scalar(Scalar::Number(n)) if !n.is_finite() => {
                result = FormulaValue::error("#NUM!", "Non-finite calculation result");
            }
            FormulaValue::Array(rows) => {
                if rows
                    .iter()
                    .any(|row| row.len() != rows.first().map_or(0, Vec::len))
                {
                    return Err(KernelError::new(
                        "FORMULA_ARRAY_SHAPE_INVALID",
                        "A function produced a nonrectangular array",
                    ));
                }
                for value in rows.iter_mut().flatten() {
                    if matches!(value,Scalar::Number(n) if !n.is_finite()) {
                        *value = Scalar::error("#NUM!", "Non-finite calculation result");
                    }
                }
            }
            _ => {}
        }
        self.validate_revision()?;
        Ok(result)
    }
    pub fn cell_output(&self, address: &CellAddress) -> KernelResult<FormulaValue> {
        self.tick()?;
        if let Some(value) = self.overrides.and_then(|values| values.get(address)) {
            return Ok(FormulaValue::Scalar(value.clone()));
        }
        if let Some(v) = self.memo.borrow().get(address) {
            return Ok(v.clone());
        }
        if !self.runtime.dirty.contains(address) && !self.forced.borrow().contains(address) {
            if let Some(v) = self.runtime.cached.get(address) {
                return Ok(v.clone());
            }
        }
        if let Some(entry) = self.runtime.formulas.get(address) {
            if !self.stack.borrow_mut().insert(address.clone()) {
                return Ok(FormulaValue::error("#REF!", "Circular formula dependency"));
            }
            let result = self.output(&entry.ast, address);
            self.stack.borrow_mut().remove(address);
            let value = result?;
            let value = self.validate_spill(address, value)?;
            self.publish_spill(address, &value)?;
            self.memo
                .borrow_mut()
                .insert(address.clone(), value.clone());
            return Ok(value);
        }
        if let Some(value) = self.runtime.values.get(address) {
            return Ok(FormulaValue::Scalar(value.clone()));
        }
        if let Some(value) = self.spill_scalar(address)? {
            return Ok(FormulaValue::Scalar(value));
        }
        let cell = self.reader.read_cell(address)?;
        if let Some(cell) = cell {
            if cell.formula.is_some() {
                return Err(KernelError::new(
                    "FORMULA_REGISTRY_MISSING",
                    "Authored formula has not entered the canonical AST index",
                )
                .at(format!(
                    "{}:{}:{}",
                    address.sheet_id, address.row, address.column
                )));
            }
            return Ok(FormulaValue::Scalar(cell.value));
        }
        Ok(FormulaValue::Scalar(Scalar::Null))
    }
    pub fn spill_scalar(&self, address: &CellAddress) -> KernelResult<Option<Scalar>> {
        let mut owners = self.runtime.spill_index.affected(address);
        owners.extend(self.spill_index.borrow().affected(address));
        for owner in owners {
            if !self.memo.borrow().contains_key(&owner)
                && (self.runtime.dirty.contains(&owner)
                    || self.forced.borrow().contains(&owner)
                    || !self.runtime.cached.contains_key(&owner))
            {
                self.cell_output(&owner)?;
            }
            let memo = self.memo.borrow();
            let value = memo.get(&owner).or_else(|| self.runtime.cached.get(&owner));
            if let Some(FormulaValue::Array(values)) = value {
                if address.row >= owner.row && address.column >= owner.column {
                    if let Some(value) = values
                        .get((address.row - owner.row) as usize)
                        .and_then(|row| row.get((address.column - owner.column) as usize))
                    {
                        return Ok(Some(value.clone()));
                    }
                }
            }
        }
        Ok(None)
    }
    fn cell_scalar(&self, address: &CellAddress) -> KernelResult<Scalar> {
        self.tick()?;
        if let Some(value) = self.overrides.and_then(|v| v.get(address)) {
            return Ok(value.clone());
        }
        if let Some(value) = self.memo.borrow().get(address) {
            return Ok(value.scalar());
        }
        if !self.runtime.dirty.contains(address) && !self.forced.borrow().contains(address) {
            if let Some(value) = self.runtime.cached.get(address) {
                return Ok(value.scalar());
            }
        }
        self.cell_output(address).map(|v| v.scalar())
    }
    fn publish_spill(&self, address: &CellAddress, value: &FormulaValue) -> KernelResult<()> {
        let (rows, cols) = value.shape();
        let range = if rows > 1 || cols > 1 {
            Some(RangeRef {
                sheet_id: address.sheet_id.clone(),
                start_row: address.row,
                end_row: address.row + rows as u32 - 1,
                start_column: address.column,
                end_column: address.column + cols as u32 - 1,
            })
        } else {
            None
        };
        let old = self
            .spill_updates
            .borrow()
            .get(address)
            .cloned()
            .unwrap_or_else(|| self.runtime.spills.get(address).cloned());
        let prior = self
            .memo
            .borrow()
            .get(address)
            .cloned()
            .or_else(|| self.runtime.cached.get(address).cloned());
        if old != range || ((old.is_some() || range.is_some()) && prior.as_ref() != Some(value)) {
            self.spill_changed.borrow_mut().insert(address.clone());
        }
        if old.is_some() || range.is_some() {
            self.spill_updates
                .borrow_mut()
                .insert(address.clone(), range.clone());
            self.spill_index.borrow_mut().remove(address);
            if let Some(range) = range {
                self.spill_index
                    .borrow_mut()
                    .add_dynamic(address, vec![range]);
            }
        }
        Ok(())
    }
    fn validate_spill(
        &self,
        address: &CellAddress,
        value: FormulaValue,
    ) -> KernelResult<FormulaValue> {
        let (rows, cols) = value.shape();
        if rows == 0 || cols == 0 {
            return Ok(FormulaValue::error("#CALC!", "Empty array result"));
        }
        if rows == 1 && cols == 1 {
            return Ok(value);
        }
        self.allocate(rows, cols)?;
        if address.row as usize + rows > MAX_ROWS as usize
            || address.column as usize + cols > MAX_COLUMNS as usize
        {
            return Ok(FormulaValue::error(
                "#SPILL!",
                "Spill extends beyond the worksheet",
            ));
        }
        let range = RangeRef {
            sheet_id: address.sheet_id.clone(),
            start_row: address.row,
            end_row: address.row + rows as u32 - 1,
            start_column: address.column,
            end_column: address.column + cols as u32 - 1,
        };
        self.record(address, &range);
        let mut blocked = false;
        self.reader.read_range(&range, &mut |a, c| {
            self.tick()?;
            if &a != address && (c.formula.is_some() || c.value != Scalar::Null) {
                blocked = true;
            }
            Ok(())
        })?;
        let lo = CellAddress {
            sheet_id: address.sheet_id.clone(),
            row: range.start_row,
            column: 0,
        };
        let hi = CellAddress {
            sheet_id: address.sheet_id.clone(),
            row: range.end_row,
            column: MAX_COLUMNS - 1,
        };
        for (a, v) in self.runtime.values.range(lo.clone()..=hi.clone()) {
            if a != address && range.contains(a) && v != &Scalar::Null {
                blocked = true;
            }
        }
        for (a, _) in self.runtime.formulas.range(lo.clone()..=hi.clone()) {
            if a != address && range.contains(a) {
                blocked = true;
            }
        }
        if let Some(overrides) = self.overrides {
            for (a, v) in overrides.range(lo..=hi) {
                if a != address && range.contains(a) && v != &Scalar::Null {
                    blocked = true;
                }
            }
        }
        let mut owners = self.runtime.spill_index.affected_range(&range);
        owners.extend(self.spill_index.borrow().affected_range(&range));
        for owner in owners {
            if &owner != address {
                let footprint = self
                    .spill_updates
                    .borrow()
                    .get(&owner)
                    .cloned()
                    .unwrap_or_else(|| self.runtime.spills.get(&owner).cloned());
                if footprint.is_some_and(|footprint| footprint.intersects(&range)) {
                    blocked = true;
                }
            }
        }
        if blocked {
            Ok(FormulaValue::error(
                "#SPILL!",
                "Array destination contains authored content",
            ))
        } else {
            Ok(value)
        }
    }
    fn record(&self, current: &CellAddress, range: &RangeRef) {
        let mut deps = self.dynamic.borrow_mut();
        let values = deps.entry(current.clone()).or_default();
        if !values.contains(range) {
            values.push(range.clone());
        }
    }
    fn reference(&self, value: &Value) -> KernelResult<RangeRef> {
        match value {
            Value::Reference(r) => Ok(r.clone()),
            _ => Err(KernelError::new("#VALUE!", "A cell reference is required")),
        }
    }
    fn scalar(&self, value: &Value) -> KernelResult<Scalar> {
        match value {
            Value::Union(values) => values
                .first()
                .ok_or_else(|| KernelError::new("#NULL!", "Empty reference union"))
                .and_then(|v| self.scalar(v)),
            Value::Missing => Ok(Scalar::Null),
            Value::Data(v) => Ok(v.scalar()),
            Value::Reference(r) => self.cell_scalar(&CellAddress {
                sheet_id: r.sheet_id.clone(),
                row: r.start_row,
                column: r.start_column,
            }),
            Value::Lambda(_) => Err(KernelError::new("#CALC!", "A LAMBDA must be invoked")),
        }
    }
    fn materialize(&self, value: &Value) -> KernelResult<FormulaValue> {
        match value {
            Value::Union(_) => Ok(FormulaValue::error(
                "#VALUE!",
                "A multi-area reference requires a reference-aware function",
            )),
            Value::Missing => Ok(FormulaValue::Scalar(Scalar::Null)),
            Value::Data(v) => Ok(v.clone()),
            Value::Lambda(_) => Ok(FormulaValue::error("#CALC!", "A LAMBDA must be invoked")),
            Value::Reference(r) => {
                let rows = (r.end_row - r.start_row + 1) as usize;
                let cols = (r.end_column - r.start_column + 1) as usize;
                self.allocate(rows, cols)?;
                if rows == 1 && cols == 1 {
                    return self
                        .cell_scalar(&CellAddress {
                            sheet_id: r.sheet_id.clone(),
                            row: r.start_row,
                            column: r.start_column,
                        })
                        .map(FormulaValue::Scalar);
                }
                let mut out = vec![vec![Scalar::Null; cols]; rows];
                for row in 0..rows {
                    for col in 0..cols {
                        out[row][col] = self.cell_scalar(&CellAddress {
                            sheet_id: r.sheet_id.clone(),
                            row: r.start_row + row as u32,
                            column: r.start_column + col as u32,
                        })?;
                    }
                }
                Ok(FormulaValue::Array(out))
            }
        }
    }
    fn shape(&self, v: &Value) -> (usize, usize) {
        match v {
            Value::Union(_) => (0, 0),
            Value::Missing => (1, 1),
            Value::Data(v) => v.shape(),
            Value::Reference(r) => (
                (r.end_row - r.start_row + 1) as usize,
                (r.end_column - r.start_column + 1) as usize,
            ),
            Value::Lambda(_) => (0, 0),
        }
    }
    fn at(&self, v: &Value, row: usize, col: usize) -> KernelResult<Scalar> {
        self.tick()?;
        match v {
            Value::Missing if row == 0 && col == 0 => Ok(Scalar::Null),
            Value::Reference(r) => self.cell_scalar(&CellAddress {
                sheet_id: r.sheet_id.clone(),
                row: r.start_row + row as u32,
                column: r.start_column + col as u32,
            }),
            Value::Data(FormulaValue::Scalar(v)) if row == 0 && col == 0 => Ok(v.clone()),
            Value::Data(FormulaValue::Array(a)) => a
                .get(row)
                .and_then(|r| r.get(col))
                .cloned()
                .ok_or_else(|| KernelError::new("#VALUE!", "Array shape mismatch")),
            _ => Err(KernelError::new("#VALUE!", "Array shape mismatch")),
        }
    }
    /// Streams authored reference values through the canonical page cursor.
    fn visit(
        &self,
        v: &Value,
        f: &mut dyn FnMut(Scalar, bool, Option<CellAddress>) -> KernelResult<()>,
    ) -> KernelResult<()> {
        match v {
            Value::Union(values) => {
                for value in values {
                    self.visit(value, f)?;
                }
                Ok(())
            }
            Value::Missing => f(Scalar::Null, false, None),
            Value::Reference(range) => {
                let mut overrides = BTreeSet::new();
                self.reader.read_range(range, &mut |address, cell| {
                    self.tick()?;
                    let value = if self.runtime.formulas.contains_key(&address)
                        || self.runtime.values.contains_key(&address)
                        || self
                            .overrides
                            .is_some_and(|values| values.contains_key(&address))
                    {
                        overrides.insert(address.clone());
                        self.cell_scalar(&address)?
                    } else if cell.formula.is_some() {
                        self.cell_scalar(&address)?
                    } else {
                        cell.value
                    };
                    f(value, true, Some(address))
                })?;
                let lo = CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row: range.start_row,
                    column: 0,
                };
                let hi = CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row: range.end_row,
                    column: MAX_COLUMNS - 1,
                };
                for (address, _) in self.runtime.values.range(lo.clone()..=hi.clone()) {
                    if range.contains(address) && !overrides.contains(address) {
                        self.tick()?;
                        f(self.cell_scalar(address)?, true, Some(address.clone()))?;
                        overrides.insert(address.clone());
                    }
                }
                for (address, _) in self.runtime.formulas.range(lo.clone()..=hi.clone()) {
                    if range.contains(address) && !overrides.contains(address) {
                        self.tick()?;
                        f(self.cell_scalar(address)?, true, Some(address.clone()))?;
                        overrides.insert(address.clone());
                    }
                }
                if let Some(values) = self.overrides {
                    for (address, value) in values.range(lo..=hi) {
                        if range.contains(address) && !overrides.contains(address) {
                            self.tick()?;
                            f(value.clone(), true, Some(address.clone()))?;
                        }
                    }
                }
                let mut roots = self.runtime.spill_index.affected_range(range);
                roots.extend(self.spill_index.borrow().affected_range(range));
                for root in roots {
                    let result = self.cell_output(&root)?;
                    let footprint = self
                        .spill_updates
                        .borrow()
                        .get(&root)
                        .cloned()
                        .unwrap_or_else(|| self.runtime.spills.get(&root).cloned());
                    if let (Some(footprint), FormulaValue::Array(values)) = (footprint, result) {
                        for row in range.start_row.max(footprint.start_row)
                            ..=range.end_row.min(footprint.end_row)
                        {
                            for column in range.start_column.max(footprint.start_column)
                                ..=range.end_column.min(footprint.end_column)
                            {
                                let address = CellAddress {
                                    sheet_id: range.sheet_id.clone(),
                                    row,
                                    column,
                                };
                                if address != root {
                                    self.tick()?;
                                    f(
                                        values[(row - footprint.start_row) as usize]
                                            [(column - footprint.start_column) as usize]
                                            .clone(),
                                        true,
                                        Some(address),
                                    )?;
                                }
                            }
                        }
                    }
                }
                Ok(())
            }
            Value::Data(FormulaValue::Scalar(v)) => {
                self.tick()?;
                f(v.clone(), false, None)
            }
            Value::Data(FormulaValue::Array(a)) => {
                for v in a.iter().flatten() {
                    self.tick()?;
                    f(v.clone(), true, None)?;
                }
                Ok(())
            }
            Value::Lambda(_) => Err(KernelError::new(
                "#VALUE!",
                "Function cannot aggregate a LAMBDA",
            )),
        }
    }
    fn eval(&self, e: &Expr, current: &CellAddress, env: &Env) -> KernelResult<Value> {
        self.tick()?;
        let depth = self.depth.get();
        if depth >= self.runtime.context.max_recursion {
            return Err(KernelError::new(
                "RESOURCE_BUDGET_EXCEEDED",
                "Formula recursion budget exhausted",
            ));
        }
        self.depth.set(depth + 1);
        let result = self.eval_inner(e, current, env);
        self.depth.set(depth);
        let result = match result {
            Err(e) if e.code.starts_with('#') => Ok(Value::error(&e.code, e.message)),
            other => other,
        };
        if self.tracing {
            if let Ok(value) = &result {
                if !matches!(value, Value::Lambda(_)) {
                    let value = self.materialize(value)?;
                    self.allocate(self.trace_steps.borrow().len() + 1, 1)?;
                    self.trace_steps
                        .borrow_mut()
                        .push(crate::FormulaEvaluationTraceStep {
                            expression: parser::format(e),
                            value,
                        });
                }
            }
        }
        result
    }
    fn eval_inner(&self, e: &Expr, current: &CellAddress, env: &Env) -> KernelResult<Value> {
        match e {
            Expr::Scalar(v) => Ok(Value::scalar(v.clone())),
            Expr::Missing => Ok(Value::Missing),
            Expr::Reference(r) => {
                let range = self.runtime.resolved_range(r)?;
                self.record(current, &range);
                Ok(Value::Reference(range))
            }
            Expr::Name(name) => {
                let key = name.to_uppercase();
                if let Some(v) = env.get(&key) {
                    return Ok(v.clone());
                }
                if let Some(expr) = self.runtime.resolve_name(&key, current) {
                    return self.eval(expr, current, env);
                }
                Err(KernelError::new("#NAME?", format!("Undefined name {name}")))
            }
            Expr::Structured(reference) => {
                let range = self.runtime.resolve_table_reference(reference, current)?;
                self.record(current, &range);
                Ok(Value::Reference(range))
            }
            Expr::Array(rows) => {
                let cols = rows.first().map_or(0, Vec::len);
                self.allocate(rows.len(), cols)?;
                if cols == 0 || rows.iter().any(|r| r.len() != cols) {
                    return Err(KernelError::new("#VALUE!", "Nonrectangular array"));
                }
                let mut out = Vec::new();
                for row in rows {
                    let mut r = Vec::new();
                    for e in row {
                        r.push(self.scalar(&self.eval(e, current, env)?)?);
                    }
                    out.push(r);
                }
                Ok(Value::Data(FormulaValue::Array(out)))
            }
            Expr::Spill(e) => {
                let v = self.eval(e, current, env)?;
                let range = self.reference(&v)?;
                if range.start_row != range.end_row || range.start_column != range.end_column {
                    return Err(KernelError::new(
                        "#REF!",
                        "Spill operator requires an anchor cell",
                    ));
                }
                let address = CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row: range.start_row,
                    column: range.start_column,
                };
                let out = self.cell_output(&address)?;
                if !matches!(out, FormulaValue::Array(_)) {
                    return Err(KernelError::new("#REF!", "Reference is not a spill anchor"));
                }
                Ok(Value::Data(out))
            }
            Expr::Unary(op, e) => {
                let v = self.eval(e, current, env)?;
                if op == "@" {
                    if let Value::Reference(r) = v {
                        let row = if r.start_row == r.end_row {
                            r.start_row
                        } else {
                            current.row
                        };
                        let column = if r.start_column == r.end_column {
                            r.start_column
                        } else {
                            current.column
                        };
                        let a = CellAddress {
                            sheet_id: r.sheet_id.clone(),
                            row,
                            column,
                        };
                        if !r.contains(&a) {
                            return Err(KernelError::new(
                                "#VALUE!",
                                "No implicit reference intersection",
                            ));
                        }
                        return Ok(Value::scalar(self.cell_scalar(&a)?));
                    }
                    return Ok(Value::scalar(self.scalar(&v)?));
                }
                let data = self.materialize(&v)?;
                let mut rows = data.matrix();
                for x in rows.iter_mut().flatten() {
                    *x = match num(x) {
                        Ok(n) => number(match op.as_str() {
                            "-" => -n,
                            "%" => n / 100.,
                            _ => n,
                        }),
                        Err(e) => Scalar::error(&e.code, e.message),
                    };
                }
                Ok(Value::Data(if matches!(data, FormulaValue::Scalar(_)) {
                    FormulaValue::Scalar(rows[0][0].clone())
                } else {
                    FormulaValue::Array(rows)
                }))
            }
            Expr::Binary(op, a, b) => {
                let a = self.eval(a, current, env)?;
                let b = self.eval(b, current, env)?;
                if op == "," {
                    if !matches!(a, Value::Reference(_) | Value::Union(_))
                        || !matches!(b, Value::Reference(_) | Value::Union(_))
                    {
                        return Err(KernelError::new(
                            "#VALUE!",
                            "Reference union operands must be references",
                        ));
                    }
                    return Ok(Value::Union(vec![a, b]));
                }
                if op == ":" {
                    let a = self.reference(&a)?;
                    let b = self.reference(&b)?;
                    if a.sheet_id != b.sheet_id {
                        return Err(KernelError::new("#REF!", "A range cannot span worksheets"));
                    }
                    return Ok(Value::Reference(RangeRef {
                        sheet_id: a.sheet_id,
                        start_row: a.start_row.min(b.start_row),
                        end_row: a.end_row.max(b.end_row),
                        start_column: a.start_column.min(b.start_column),
                        end_column: a.end_column.max(b.end_column),
                    }));
                }
                let (ar, ac) = self.shape(&a);
                let (br, bc) = self.shape(&b);
                let rows = ar.max(br);
                let cols = ac.max(bc);
                if (ar != br && ar != 1 && br != 1) || (ac != bc && ac != 1 && bc != 1) {
                    return Err(KernelError::new("#N/A", "Arrays cannot be broadcast"));
                }
                self.allocate(rows, cols)?;
                let mut out = vec![vec![Scalar::Null; cols]; rows];
                for row in 0..rows {
                    for col in 0..cols {
                        out[row][col] = binary(
                            op,
                            &self.at(
                                &a,
                                if ar == 1 { 0 } else { row },
                                if ac == 1 { 0 } else { col },
                            )?,
                            &self.at(
                                &b,
                                if br == 1 { 0 } else { row },
                                if bc == 1 { 0 } else { col },
                            )?,
                        );
                    }
                }
                Ok(Value::Data(if rows == 1 && cols == 1 {
                    FormulaValue::Scalar(out[0][0].clone())
                } else {
                    FormulaValue::Array(out)
                }))
            }
            Expr::Call(name, args) => self.call(name, args, current, env),
            Expr::Invoke(callee, args) => {
                let callee = self.eval(callee, current, env)?;
                let values = args
                    .iter()
                    .map(|a| self.eval(a, current, env))
                    .collect::<KernelResult<Vec<_>>>()?;
                self.invoke(&callee, values, current)
            }
        }
    }
    fn invoke(
        &self,
        callee: &Value,
        args: Vec<Value>,
        current: &CellAddress,
    ) -> KernelResult<Value> {
        let Value::Lambda(lambda) = callee else {
            return Err(KernelError::new("#VALUE!", "Value is not callable"));
        };
        if args.len() != lambda.params.len() {
            return Err(KernelError::new(
                "#VALUE!",
                "LAMBDA parameter count mismatch",
            ));
        }
        let mut env = lambda.env.clone();
        if let Some(name) = &lambda.recursive_name {
            env.insert(name.clone(), callee.clone());
        }
        for (name, value) in lambda.params.iter().zip(args) {
            env.insert(name.clone(), value);
        }
        self.eval(&lambda.body, current, &env)
    }
    fn call(
        &self,
        name: &str,
        args: &[Expr],
        current: &CellAddress,
        env: &Env,
    ) -> KernelResult<Value> {
        let n = name.to_uppercase();
        let name = n.as_str();
        let ev = |i: usize| -> KernelResult<Value> {
            args.get(i)
                .map(|a| self.eval(a, current, env))
                .unwrap_or_else(|| Ok(Value::scalar(Scalar::Null)))
        };
        let scalar = |i: usize| self.scalar(&ev(i)?);
        let number = |i: usize| num(&scalar(i)?);
        match name {
            "SJS.TABLE" => {
                if args.len() != 3 && args.len() != 5 {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "SJS.TABLE requires a result cell and one or two value/input-cell pairs",
                    ));
                }
                let result_range = self.reference(&ev(0)?)?;
                if result_range.start_row != result_range.end_row
                    || result_range.start_column != result_range.end_column
                {
                    return Err(KernelError::new("#VALUE!", "SJS.TABLE result must be one cell"));
                }
                let result_address = CellAddress {
                    sheet_id: result_range.sheet_id,
                    row: result_range.start_row,
                    column: result_range.start_column,
                };
                if &result_address == current {
                    return Err(KernelError::new("#REF!", "SJS.TABLE result cannot reference itself"));
                }
                let input_address = |index: usize| -> KernelResult<CellAddress> {
                    let range = self.reference(&ev(index)?)?;
                    if range.start_row != range.end_row || range.start_column != range.end_column {
                        return Err(KernelError::new("#VALUE!", "SJS.TABLE input must be one cell"));
                    }
                    Ok(CellAddress {
                        sheet_id: range.sheet_id,
                        row: range.start_row,
                        column: range.start_column,
                    })
                };
                let first_values = self.materialize(&ev(1)?)?.matrix();
                let first_input = input_address(2)?;
                let calculate = |overrides: &BTreeMap<CellAddress, Scalar>| -> KernelResult<Scalar> {
                    let session = Session::with_overrides(
                        self.runtime,
                        self.reader,
                        self.services,
                        overrides,
                    )?;
                    Ok(session.cell_output(&result_address)?.scalar())
                };
                let output = if args.len() == 3 {
                    self.allocate(first_values.len(), first_values.first().map_or(0, Vec::len))?;
                    let mut output = Vec::with_capacity(first_values.len());
                    for row in first_values {
                        let mut output_row = Vec::with_capacity(row.len());
                        for value in row {
                            output_row.push(calculate(&BTreeMap::from([(first_input.clone(), value)]))?);
                        }
                        output.push(output_row);
                    }
                    output
                } else {
                    let second_values = self.materialize(&ev(3)?)?.matrix();
                    let second_input = input_address(4)?;
                    let row_values: Vec<_> = first_values.into_iter().flatten().collect();
                    let column_values: Vec<_> = second_values.into_iter().flatten().collect();
                    self.allocate(row_values.len(), column_values.len())?;
                    let mut output = Vec::with_capacity(row_values.len());
                    for row_value in row_values {
                        let mut output_row = Vec::with_capacity(column_values.len());
                        for column_value in &column_values {
                            output_row.push(calculate(&BTreeMap::from([
                                (first_input.clone(), row_value.clone()),
                                (second_input.clone(), column_value.clone()),
                            ]))?);
                        }
                        output.push(output_row);
                    }
                    output
                };
                return Ok(Value::Data(FormulaValue::Array(output)));
            }
            "INDEX" => {
                arity(args, 2, 4)?;
                let input = ev(0)?;
                let area = if args.len() > 3 {
                    number(3)?.trunc() as i64
                } else {
                    1
                };
                let input = if let Value::Union(values) = input {
                    fn flatten(values: Vec<Value>, out: &mut Vec<Value>) {
                        for value in values {
                            if let Value::Union(values) = value {
                                flatten(values, out)
                            } else {
                                out.push(value)
                            }
                        }
                    }
                    let mut areas = Vec::new();
                    flatten(values, &mut areas);
                    if area < 1 || area as usize > areas.len() {
                        return Err(KernelError::new(
                            "#REF!",
                            "INDEX area number is outside reference union",
                        ));
                    }
                    areas.remove(area as usize - 1)
                } else {
                    if area != 1 {
                        return Err(KernelError::new(
                            "#REF!",
                            "INDEX area number is outside reference",
                        ));
                    }
                    input
                };
                if let Value::Reference(mut range) = input {
                    let mut row = number(1)?.trunc() as i64;
                    let mut col = if args.len() > 2 {
                        number(2)?.trunc() as i64
                    } else {
                        1
                    };
                    let rows = (range.end_row - range.start_row + 1) as i64;
                    let cols = (range.end_column - range.start_column + 1) as i64;
                    if rows == 1 && args.len() == 2 {
                        col = row;
                        row = 1;
                    }
                    if row < 0 || col < 0 || row > rows || col > cols {
                        return Err(KernelError::new(
                            "#REF!",
                            "INDEX row or column is outside reference",
                        ));
                    }
                    if row > 0 {
                        range.start_row += row as u32 - 1;
                        range.end_row = range.start_row;
                    }
                    if col > 0 {
                        range.start_column += col as u32 - 1;
                        range.end_column = range.start_column;
                    }
                    self.record(current, &range);
                    return Ok(Value::Reference(range));
                }
                let mut values = vec![self.materialize(&input)?];
                for i in 1..args.len().min(3) {
                    values.push(self.materialize(&ev(i)?)?);
                }
                return array_functions::call_with_limit(
                    name,
                    &values,
                    self.runtime.context.max_array_cells,
                )
                .expect("INDEX is registered")
                .map(Value::Data);
            }
            "LET" => {
                if args.len() < 3 || args.len() % 2 != 1 {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "LET requires name/value pairs and a result",
                    ));
                }
                let mut scope = env.clone();
                for pair in args[..args.len() - 1].chunks_exact(2) {
                    let Expr::Name(name) = &pair[0] else {
                        return Err(KernelError::new("#VALUE!", "Invalid LET variable name"));
                    };
                    let key = name.to_uppercase();
                    let mut value = self.eval(&pair[1], current, &scope)?;
                    if let Value::Lambda(lambda) = &value {
                        let mut lambda = (**lambda).clone();
                        lambda.recursive_name = Some(key.clone());
                        value = Value::Lambda(Rc::new(lambda));
                    }
                    scope.insert(key, value);
                }
                return self.eval(args.last().unwrap(), current, &scope);
            }
            "LAMBDA" => {
                if args.is_empty() {
                    return Err(KernelError::new("#VALUE!", "LAMBDA requires a body"));
                }
                let mut params = Vec::new();
                for arg in &args[..args.len() - 1] {
                    let Expr::Name(name) = arg else {
                        return Err(KernelError::new("#VALUE!", "Invalid LAMBDA parameter"));
                    };
                    let key = name.to_uppercase();
                    if params.contains(&key) {
                        return Err(KernelError::new("#VALUE!", "Duplicate LAMBDA parameter"));
                    }
                    params.push(key);
                }
                return Ok(Value::Lambda(Rc::new(Lambda {
                    params,
                    body: args.last().unwrap().clone(),
                    env: env.clone(),
                    recursive_name: None,
                })));
            }
            "ISOMITTED" => {
                arity(args, 1, 1)?;
                return Ok(Value::scalar(Scalar::Boolean(matches!(
                    ev(0)?,
                    Value::Missing
                ))));
            }
            "IF" => {
                arity(args, 2, 3)?;
                let test = ev(0)?;
                let (rows, cols) = self.shape(&test);
                if rows == 1 && cols == 1 {
                    return if truth(&self.scalar(&test)?)? {
                        ev(1)
                    } else if args.len() == 3 {
                        ev(2)
                    } else {
                        Ok(Value::scalar(Scalar::Boolean(false)))
                    };
                }
                self.allocate(rows, cols)?;
                let mut yes = None;
                let mut no = None;
                let mut out = vec![vec![Scalar::Null; cols]; rows];
                for r in 0..rows {
                    for c in 0..cols {
                        let condition = self.at(&test, r, c)?;
                        out[r][c] = match truth(&condition) {
                            Err(e) => Scalar::error(&e.code, e.message),
                            Ok(t) => {
                                let slot = if t { &mut yes } else { &mut no };
                                if slot.is_none() {
                                    *slot = Some(if t {
                                        ev(1)?
                                    } else if args.len() == 3 {
                                        ev(2)?
                                    } else {
                                        Value::scalar(Scalar::Boolean(false))
                                    });
                                }
                                let v = slot.as_ref().unwrap();
                                let (vr, vc) = self.shape(v);
                                self.at(
                                    v,
                                    if vr == 1 { 0 } else { r },
                                    if vc == 1 { 0 } else { c },
                                )?
                            }
                        };
                    }
                }
                return Ok(Value::Data(FormulaValue::Array(out)));
            }
            "IFERROR" | "IFNA" => {
                arity(args, 2, 2)?;
                let value = self.materialize(&ev(0)?)?;
                let mut fallback = None;
                let mut rows = value.matrix();
                for (r, row) in rows.iter_mut().enumerate() {
                    for (c, v) in row.iter_mut().enumerate() {
                        if matches!(v,Scalar::Error(e) if name=="IFERROR"||e.code=="#N/A") {
                            if fallback.is_none() {
                                fallback = Some(ev(1)?);
                            }
                            let f = fallback.as_ref().unwrap();
                            let (fr, fc) = self.shape(f);
                            *v = self.at(
                                f,
                                if fr == 1 { 0 } else { r },
                                if fc == 1 { 0 } else { c },
                            )?;
                        }
                    }
                }
                return Ok(Value::Data(if matches!(value, FormulaValue::Scalar(_)) {
                    FormulaValue::Scalar(rows[0][0].clone())
                } else {
                    FormulaValue::Array(rows)
                }));
            }
            "IFS" => {
                if args.is_empty() || args.len() % 2 != 0 {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "IFS requires condition/value pairs",
                    ));
                }
                for i in (0..args.len()).step_by(2) {
                    if truth(&scalar(i)?)? {
                        return ev(i + 1);
                    }
                }
                return Err(KernelError::new("#N/A", "No IFS condition matched"));
            }
            "SWITCH" => {
                if args.len() < 3 {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "SWITCH requires expression and case/value pairs",
                    ));
                }
                let value = scalar(0)?;
                for i in (1..args.len() - 1).step_by(2) {
                    if compare(&value, &scalar(i)?).is_eq() {
                        return ev(i + 1);
                    }
                }
                return if args.len() % 2 == 0 {
                    ev(args.len() - 1)
                } else {
                    Err(KernelError::new("#N/A", "No SWITCH case matched"))
                };
            }
            "CHOOSE" => {
                arity(args, 2, 255)?;
                let index = number(0)?.trunc();
                if index < 1. || index >= args.len() as f64 {
                    return Err(KernelError::new("#VALUE!", "CHOOSE index outside choices"));
                }
                return ev(index as usize);
            }
            "MAP" | "REDUCE" | "SCAN" | "BYROW" | "BYCOL" | "MAKEARRAY" => {
                return self.lambda_array(name, args, current, env);
            }
            "ROW" | "COLUMN" => {
                arity(args, 0, 1)?;
                let range = if args.is_empty() {
                    RangeRef {
                        sheet_id: current.sheet_id.clone(),
                        start_row: current.row,
                        end_row: current.row,
                        start_column: current.column,
                        end_column: current.column,
                    }
                } else {
                    self.reference(&ev(0)?)?
                };
                let row = name == "ROW";
                let start = if row {
                    range.start_row
                } else {
                    range.start_column
                };
                let end = if row { range.end_row } else { range.end_column };
                let values: Vec<Scalar> = (start..=end)
                    .map(|i| Scalar::Number((i + 1) as f64))
                    .collect();
                return Ok(Value::Data(if start == end {
                    FormulaValue::Scalar(values[0].clone())
                } else if row {
                    FormulaValue::Array(values.into_iter().map(|v| vec![v]).collect())
                } else {
                    FormulaValue::Array(vec![values])
                }));
            }
            "ADDRESS" => {
                arity(args, 2, 5)?;
                let row = number(0)?.trunc() as i64;
                let col = number(1)?.trunc() as i64;
                let mode = if args.len() > 2 && !matches!(args[2], Expr::Missing) {
                    number(2)?.trunc() as i64
                } else {
                    1
                };
                let a1 = if args.len() > 3 && !matches!(args[3], Expr::Missing) {
                    truth(&scalar(3)?)?
                } else {
                    true
                };
                if row < 1
                    || col < 1
                    || row > MAX_ROWS as i64
                    || col > MAX_COLUMNS as i64
                    || !(1..=4).contains(&mode)
                {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "ADDRESS row/column/mode outside valid bounds",
                    ));
                }
                let mut result = if a1 {
                    format!(
                        "{}{}{}{}",
                        if mode == 1 || mode == 3 { "$" } else { "" },
                        column_name(col as u32),
                        if mode == 1 || mode == 2 { "$" } else { "" },
                        row
                    )
                } else {
                    format!(
                        "R{}C{}",
                        if mode == 1 || mode == 2 {
                            row.to_string()
                        } else {
                            format!("[{row}]")
                        },
                        if mode == 1 || mode == 3 {
                            col.to_string()
                        } else {
                            format!("[{col}]")
                        }
                    )
                };
                if args.len() > 4 {
                    let sheet = text(&scalar(4)?);
                    if !sheet.is_empty() {
                        result = format!("'{}'!{result}", sheet.replace('\'', "''"));
                    }
                }
                return Ok(Value::scalar(Scalar::Text(result)));
            }
            "OFFSET" => {
                arity(args, 3, 5)?;
                let base = self.reference(&ev(0)?)?;
                let row = base.start_row as i64 + number(1)?.trunc() as i64;
                let col = base.start_column as i64 + number(2)?.trunc() as i64;
                let height = if args.len() > 3 && !matches!(args[3], Expr::Missing) {
                    number(3)?.trunc() as i64
                } else {
                    (base.end_row - base.start_row + 1) as i64
                };
                let width = if args.len() > 4 && !matches!(args[4], Expr::Missing) {
                    number(4)?.trunc() as i64
                } else {
                    (base.end_column - base.start_column + 1) as i64
                };
                if row < 0
                    || col < 0
                    || height <= 0
                    || width <= 0
                    || row.saturating_add(height) > MAX_ROWS as i64
                    || col.saturating_add(width) > MAX_COLUMNS as i64
                {
                    return Err(KernelError::new(
                        "#REF!",
                        "OFFSET produces an invalid worksheet reference",
                    ));
                }
                let range = RangeRef {
                    sheet_id: base.sheet_id,
                    start_row: row as u32,
                    end_row: (row + height - 1) as u32,
                    start_column: col as u32,
                    end_column: (col + width - 1) as u32,
                };
                self.record(current, &range);
                return Ok(Value::Reference(range));
            }
            "INDIRECT" => {
                arity(args, 1, 2)?;
                let source = text(&scalar(0)?);
                let a1 = if args.len() > 1 {
                    truth(&scalar(1)?)?
                } else {
                    true
                };
                if let Some(expr) = self.runtime.resolve_name(&source, current) {
                    return self.eval(expr, current, env);
                }
                let range = parser::parse_reference(&source, current, a1)
                    .map_err(|e| KernelError::new("#REF!", e.message))?;
                let range = self.runtime.resolved_range(&range)?;
                self.record(current, &range);
                return Ok(Value::Reference(range));
            }
            "NOW" | "TODAY" => {
                arity(args, 0, 0)?;
                let now = self.runtime.context.now_serial.ok_or_else(|| {
                    KernelError::new(
                        "CALCULATION_CONTEXT_REQUIRED",
                        "NOW/TODAY require a transaction reference instant",
                    )
                })?;
                return Ok(Value::scalar(Scalar::Number(if name == "TODAY" {
                    now.floor()
                } else {
                    now
                })));
            }
            "RAND" => {
                arity(args, 0, 0)?;
                return Ok(Value::scalar(Scalar::Number(self.random(current))));
            }
            "RANDBETWEEN" => {
                arity(args, 2, 2)?;
                let lo = number(0)?.ceil();
                let hi = number(1)?.floor();
                if lo > hi {
                    return Err(KernelError::new(
                        "#NUM!",
                        "RANDBETWEEN minimum exceeds maximum",
                    ));
                }
                return Ok(Value::scalar(super::evaluator::number(
                    lo + (self.random(current) * (hi - lo + 1.)).floor(),
                )));
            }
            "RANDARRAY" => {
                arity(args, 0, 5)?;
                let rows = if args.is_empty() {
                    1
                } else {
                    positive_dimension(number(0)?)?
                };
                let cols = if args.len() < 2 {
                    1
                } else {
                    positive_dimension(number(1)?)?
                };
                let lo = if args.len() < 3 { 0. } else { number(2)? };
                let hi = if args.len() < 4 { 1. } else { number(3)? };
                let integer = if args.len() < 5 {
                    false
                } else {
                    truth(&scalar(4)?)?
                };
                if hi < lo {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "RANDARRAY minimum exceeds maximum",
                    ));
                }
                self.allocate(rows, cols)?;
                let mut out = vec![vec![Scalar::Null; cols]; rows];
                for row in &mut out {
                    for value in row {
                        self.tick()?;
                        *value = Scalar::Number(if integer {
                            let lo = lo.ceil();
                            let hi = hi.floor();
                            if lo > hi {
                                return Err(KernelError::new(
                                    "#VALUE!",
                                    "RANDARRAY has no integer in bounds",
                                ));
                            }
                            lo + (self.random(current) * (hi - lo + 1.)).floor()
                        } else {
                            lo + self.random(current) * (hi - lo)
                        });
                    }
                }
                return Ok(Value::Data(FormulaValue::Array(out)));
            }
            _ => {}
        }
        if let Some(callee) = env.get(&n) {
            let values = args
                .iter()
                .map(|a| self.eval(a, current, env))
                .collect::<KernelResult<Vec<_>>>()?;
            return self.invoke(callee, values, current);
        }
        if let Some(expr) = self.runtime.resolve_name(&n, current) {
            let callee = self.eval(expr, current, env)?;
            let values = args
                .iter()
                .map(|a| self.eval(a, current, env))
                .collect::<KernelResult<Vec<_>>>()?;
            return self.invoke(&callee, values, current);
        }
        if STREAMING.contains(&name) {
            let values = args
                .iter()
                .map(|a| self.eval(a, current, env))
                .collect::<KernelResult<Vec<_>>>()?;
            return self.aggregate(name, &values);
        }
        // The not-found branch is the only lazy XLOOKUP argument. All match and
        // search arguments participate in the real lookup implementation.
        if name == "XLOOKUP" {
            arity(args, 3, 6)?;
            let mut values = Vec::new();
            for (i, arg) in args.iter().enumerate() {
                values.push(if i == 3 {
                    FormulaValue::error("#N/A", "Lookup value not found")
                } else {
                    self.materialize(&self.eval(arg, current, env)?)?
                });
            }
            return match array_functions::xlookup_outcome(&values) {
                Ok(Some(value)) => Ok(Value::Data(value)),
                Ok(None) if args.len() > 3 => ev(3),
                Ok(None) => Ok(Value::error("#N/A", "Lookup value not found")),
                Err(value) => Ok(Value::Data(value)),
            };
        }
        let values = args
            .iter()
            .map(|a| self.materialize(&self.eval(a, current, env)?))
            .collect::<KernelResult<Vec<_>>>()?;
        let result = if let Some(r) =
            array_functions::call_with_limit(name, &values, self.runtime.context.max_array_cells)
        {
            r
        } else if let Some(r) =
            date_stat_functions::call_with_date_system(name, &values, self.runtime.context.date1904)
        {
            r
        } else if let Some(r) = scalar_functions::call(name, &values) {
            r
        } else {
            self.services.external(name, &values)
        };
        result.map(Value::Data)
    }
    fn random(&self, current: &CellAddress) -> f64 {
        let mut counts = self.random_counts.borrow_mut();
        let count = counts.entry(current.clone()).or_default();
        let ordinal = *count;
        *count += 1;
        let mut identity = 0xcbf29ce484222325u64;
        for byte in current
            .sheet_id
            .as_bytes()
            .iter()
            .copied()
            .chain(current.row.to_le_bytes())
            .chain(current.column.to_le_bytes())
        {
            identity = (identity ^ byte as u64).wrapping_mul(0x100000001b3);
        }
        let mut x =
            self.runtime.context.random_seed ^ identity ^ ordinal.wrapping_mul(0x9e3779b97f4a7c15);
        x = (x ^ (x >> 30)).wrapping_mul(0xbf58476d1ce4e5b9);
        x = (x ^ (x >> 27)).wrapping_mul(0x94d049bb133111eb);
        ((x ^ (x >> 31)) >> 11) as f64 / (1u64 << 53) as f64
    }
    fn lambda_array(
        &self,
        name: &str,
        args: &[Expr],
        current: &CellAddress,
        env: &Env,
    ) -> KernelResult<Value> {
        let ev = |i: usize| self.eval(&args[i], current, env);
        let make = |rows: Vec<Vec<Scalar>>| Ok(Value::Data(FormulaValue::Array(rows)));
        match name {
            "MAP" => {
                arity(args, 2, 255)?;
                let lambda = ev(args.len() - 1)?;
                let inputs = args[..args.len() - 1]
                    .iter()
                    .map(|a| self.eval(a, current, env))
                    .collect::<KernelResult<Vec<_>>>()?;
                let shape = self.shape(&inputs[0]);
                if inputs.iter().any(|v| self.shape(v) != shape) {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "MAP arrays must have identical shapes",
                    ));
                }
                self.allocate(shape.0, shape.1)?;
                let mut out = vec![vec![Scalar::Null; shape.1]; shape.0];
                for r in 0..shape.0 {
                    for c in 0..shape.1 {
                        let params = inputs
                            .iter()
                            .map(|v| self.at(v, r, c).map(Value::scalar))
                            .collect::<KernelResult<Vec<_>>>()?;
                        let value = self.invoke(&lambda, params, current)?;
                        if self.shape(&value) != (1, 1) {
                            return Err(KernelError::new(
                                "#CALC!",
                                "MAP cannot return nested arrays",
                            ));
                        }
                        out[r][c] = self.scalar(&value)?;
                    }
                }
                make(out)
            }
            "REDUCE" | "SCAN" => {
                arity(args, 3, 3)?;
                let mut accumulator = ev(0)?;
                let input = ev(1)?;
                let lambda = ev(2)?;
                let (rows, cols) = self.shape(&input);
                if name == "SCAN" {
                    self.allocate(rows, cols)?;
                }
                let mut out = if name == "SCAN" {
                    vec![vec![Scalar::Null; cols]; rows]
                } else {
                    Vec::new()
                };
                for r in 0..rows {
                    for c in 0..cols {
                        accumulator = self.invoke(
                            &lambda,
                            vec![accumulator, Value::scalar(self.at(&input, r, c)?)],
                            current,
                        )?;
                        if name == "SCAN" {
                            if self.shape(&accumulator) != (1, 1) {
                                return Err(KernelError::new(
                                    "#CALC!",
                                    "SCAN cannot return nested arrays",
                                ));
                            }
                            out[r][c] = self.scalar(&accumulator)?;
                        }
                    }
                }
                if name == "SCAN" {
                    make(out)
                } else {
                    Ok(accumulator)
                }
            }
            "BYROW" | "BYCOL" => {
                arity(args, 2, 2)?;
                let input = ev(0)?;
                let lambda = ev(1)?;
                let (rows, cols) = self.shape(&input);
                let byrow = name == "BYROW";
                let count = if byrow { rows } else { cols };
                self.allocate(count, 1)?;
                let mut result = Vec::new();
                for i in 0..count {
                    let mut slice = Vec::new();
                    if byrow {
                        let mut row = Vec::new();
                        for c in 0..cols {
                            row.push(self.at(&input, i, c)?);
                        }
                        slice.push(row);
                    } else {
                        for r in 0..rows {
                            slice.push(vec![self.at(&input, r, i)?]);
                        }
                    }
                    let value = self.invoke(
                        &lambda,
                        vec![Value::Data(FormulaValue::Array(slice))],
                        current,
                    )?;
                    if self.shape(&value) != (1, 1) {
                        return Err(KernelError::new(
                            "#CALC!",
                            "BYROW/BYCOL must return one result per group",
                        ));
                    }
                    result.push(self.scalar(&value)?);
                }
                make(if byrow {
                    result.into_iter().map(|v| vec![v]).collect()
                } else {
                    vec![result]
                })
            }
            "MAKEARRAY" => {
                arity(args, 3, 3)?;
                let rows = positive_dimension(num(&self.scalar(&ev(0)?)?)?)?;
                let cols = positive_dimension(num(&self.scalar(&ev(1)?)?)?)?;
                self.allocate(rows, cols)?;
                let lambda = ev(2)?;
                let mut out = vec![vec![Scalar::Null; cols]; rows];
                for r in 0..rows {
                    for c in 0..cols {
                        let result = self.invoke(
                            &lambda,
                            vec![
                                Value::scalar(Scalar::Number((r + 1) as f64)),
                                Value::scalar(Scalar::Number((c + 1) as f64)),
                            ],
                            current,
                        )?;
                        if self.shape(&result) != (1, 1) {
                            return Err(KernelError::new(
                                "#CALC!",
                                "MAKEARRAY cannot return nested arrays",
                            ));
                        }
                        out[r][c] = self.scalar(&result)?;
                    }
                }
                make(out)
            }
            _ => unreachable!(),
        }
    }
    fn aggregate(&self, name: &str, args: &[Value]) -> KernelResult<Value> {
        if args.is_empty() {
            return Err(KernelError::new("#VALUE!", "Aggregate requires arguments"));
        }
        if matches!(
            name,
            "SUMIF"
                | "SUMIFS"
                | "COUNTIF"
                | "COUNTIFS"
                | "AVERAGEIF"
                | "AVERAGEIFS"
                | "MAXIFS"
                | "MINIFS"
        ) {
            return self.criteria_aggregate(name, args);
        }
        if name == "SUMPRODUCT" {
            let shape = self.shape(&args[0]);
            if args.iter().any(|v| self.shape(v) != shape) {
                return Err(KernelError::new(
                    "#VALUE!",
                    "SUMPRODUCT array shapes differ",
                ));
            }
            let mut sum = 0.;
            for row in 0..shape.0 {
                for col in 0..shape.1 {
                    let mut product = 1.;
                    for arg in args {
                        let v = self.at(arg, row, col)?;
                        product *= match v {
                            Scalar::Number(n) => n,
                            Scalar::Error(e) => return Err(KernelError::new(e.code, e.message)),
                            _ => 0.,
                        };
                    }
                    sum += product;
                }
            }
            return Ok(Value::scalar(number(sum)));
        }
        if name == "COUNTBLANK" {
            if args.len() != 1 {
                return Err(KernelError::new("#VALUE!", "COUNTBLANK requires one range"));
            }
            let shape = self.shape(&args[0]);
            let mut nonblank = 0u64;
            self.visit(&args[0], &mut |v, _, _| {
                if !matches!(v, Scalar::Null) && !matches!(v,Scalar::Text(t) if t.is_empty()) {
                    nonblank += 1;
                }
                Ok(())
            })?;
            return Ok(Value::scalar(Scalar::Number(
                (shape.0 as u64 * shape.1 as u64 - nonblank) as f64,
            )));
        }
        let mut operation = name;
        let mut start = 0;
        let mut visibility = None;
        let mut ignore_errors = false;
        let mut ignore_nested = false;
        if name == "SUBTOTAL" || name == "AGGREGATE" {
            let min = if name == "SUBTOTAL" { 2 } else { 3 };
            if args.len() < min {
                return Err(KernelError::new(
                    "#VALUE!",
                    "Aggregate selector and references are required",
                ));
            }
            let code = num(&self.scalar(&args[0])?)?.trunc() as i32;
            let function = if name == "SUBTOTAL" {
                if !(1..=11).contains(&code) && !(101..=111).contains(&code) {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "Invalid SUBTOTAL function number",
                    ));
                }
                visibility = Some(code >= 100);
                ignore_nested = true;
                start = 1;
                code % 100
            } else {
                if !(1..=19).contains(&code) {
                    return Err(KernelError::new(
                        "#VALUE!",
                        "Invalid AGGREGATE function number",
                    ));
                }
                let options = num(&self.scalar(&args[1])?)?.trunc() as i32;
                if !(0..=7).contains(&options) {
                    return Err(KernelError::new("#VALUE!", "Invalid AGGREGATE options"));
                }
                visibility = Some(matches!(options, 1 | 3 | 5 | 7));
                ignore_errors = matches!(options, 2 | 3 | 6 | 7);
                ignore_nested = options <= 3;
                start = 2;
                code
            };
            operation = match function {
                1 => "AVERAGE",
                2 => "COUNT",
                3 => "COUNTA",
                4 => "MAX",
                5 => "MIN",
                6 => "PRODUCT",
                7 => "STDEV",
                8 => "STDEVP",
                9 => "SUM",
                10 => "VAR",
                11 => "VARP",
                12 => "MEDIAN",
                13 => "MODE",
                14 => "LARGE",
                15 => "SMALL",
                16 => "PERCENTILE.INC",
                17 => "QUARTILE.INC",
                18 => "PERCENTILE.EXC",
                19 => "QUARTILE.EXC",
                _ => unreachable!(),
            };
        }
        let selection = matches!(
            operation,
            "MEDIAN"
                | "MODE"
                | "LARGE"
                | "SMALL"
                | "PERCENTILE.INC"
                | "PERCENTILE.EXC"
                | "QUARTILE.INC"
                | "QUARTILE.EXC"
        );
        let has_k = selection && operation != "MEDIAN" && operation != "MODE";
        let end = if has_k {
            if args.len() != start + 2 {
                return Err(KernelError::new(
                    "#VALUE!",
                    "AGGREGATE selection requires array and k",
                ));
            }
            args.len() - 1
        } else {
            args.len()
        };
        let mut acc = Accumulator::default();
        let mut logical = operation == "AND";
        let mut logical_count = 0;
        let mut selected = Vec::new();
        for arg in &args[start..end] {
            self.visit(arg, &mut |value, reference, address| {
                if let Some(manual) = visibility {
                    if let Some(address) = address {
                        let v = self.services.visibility(&address)?;
                        if v.filter_hidden || (manual && (v.manual_hidden || v.outline_hidden)) {
                            return Ok(());
                        }
                        if ignore_nested {
                            let authored = if let Some(entry) = self.runtime.formulas.get(&address)
                            {
                                Some(entry.source.clone())
                            } else {
                                self.reader
                                    .read_cell(&address)?
                                    .and_then(|cell| cell.formula)
                            };
                            if let Some(source) = authored {
                                let source =
                                    source.trim_start_matches('=').trim_start().to_uppercase();
                                if source.starts_with("SUBTOTAL(")
                                    || source.starts_with("AGGREGATE(")
                                {
                                    return Ok(());
                                }
                            }
                        }
                    }
                }
                if let Scalar::Error(e) = &value {
                    if ignore_errors || operation == "COUNT" {
                        return Ok(());
                    }
                    if operation == "COUNTA" {
                        acc.counta += 1;
                        return Ok(());
                    }
                    return Err(KernelError::new(&e.code, &e.message));
                }
                if operation == "AND" || operation == "OR" {
                    if reference && matches!(value, Scalar::Null | Scalar::Text(_)) {
                        return Ok(());
                    }
                    let v = truth(&value)?;
                    logical_count += 1;
                    if operation == "AND" {
                        logical &= v
                    } else {
                        logical |= v
                    }
                    return Ok(());
                }
                if value != Scalar::Null {
                    acc.counta += 1;
                }
                let n = match &value {
                    Scalar::Number(n) => Some(*n),
                    Scalar::Boolean(_) if !reference => Some(num(&value)?),
                    Scalar::Text(_) if !reference => match num(&value) {
                        Ok(n) => Some(n),
                        Err(_) if operation == "COUNT" => None,
                        Err(e) => return Err(e),
                    },
                    Scalar::Null if !reference => Some(0.),
                    _ => None,
                };
                if let Some(n) = n {
                    acc.push(n);
                    if selection {
                        self.allocate(selected.len() + 1, 1)?;
                        selected.push(n);
                    }
                }
                Ok(())
            })?;
        }
        if operation == "AND" || operation == "OR" {
            if logical_count == 0 {
                return Err(KernelError::new(
                    "#VALUE!",
                    "No logical values in arguments",
                ));
            }
            return Ok(Value::scalar(Scalar::Boolean(logical)));
        }
        if selection {
            selected.sort_by(f64::total_cmp);
            if selected.is_empty() {
                return Err(KernelError::new(
                    "#NUM!",
                    "No numbers for aggregate selection",
                ));
            }
            let k = if has_k {
                num(&self.scalar(args.last().unwrap())?)?
            } else {
                0.
            };
            let n = selected.len();
            let value = match operation {
                "MEDIAN" => {
                    if n % 2 == 0 {
                        (selected[n / 2 - 1] + selected[n / 2]) / 2.
                    } else {
                        selected[n / 2]
                    }
                }
                "MODE" => {
                    let mut best = selected[0];
                    let mut count = 1;
                    let mut best_count = 1;
                    for i in 1..n {
                        count = if selected[i] == selected[i - 1] {
                            count + 1
                        } else {
                            1
                        };
                        if count > best_count {
                            best_count = count;
                            best = selected[i];
                        }
                    }
                    if best_count == 1 {
                        return Err(KernelError::new("#N/A", "No repeated value"));
                    }
                    best
                }
                "LARGE" | "SMALL" => {
                    let k = k.ceil() as usize;
                    if k == 0 || k > n {
                        return Err(KernelError::new("#NUM!", "Selection index outside data"));
                    }
                    selected[if operation == "LARGE" { n - k } else { k - 1 }]
                }
                _ => {
                    let fraction = if operation.starts_with("QUARTILE") {
                        k.trunc() / 4.
                    } else {
                        k
                    };
                    let exclusive = operation.ends_with("EXC");
                    let position = if exclusive {
                        fraction * (n + 1) as f64 - 1.
                    } else {
                        fraction * (n - 1) as f64
                    };
                    if !(0. ..=1.).contains(&fraction) || position < 0. || position > (n - 1) as f64
                    {
                        return Err(KernelError::new(
                            "#NUM!",
                            "Percentile outside supported range",
                        ));
                    }
                    let lo = position.floor() as usize;
                    let hi = position.ceil() as usize;
                    selected[lo] + (selected[hi] - selected[lo]) * (position - lo as f64)
                }
            };
            return Ok(Value::scalar(number(value)));
        }
        let value = match operation {
            "COUNT" => acc.count as f64,
            "COUNTA" => acc.counta as f64,
            "SUM" => acc.sum,
            "PRODUCT" => {
                if acc.count == 0 {
                    0.
                } else {
                    acc.product
                }
            }
            "MIN" => {
                if acc.count == 0 {
                    0.
                } else {
                    acc.min
                }
            }
            "MAX" => {
                if acc.count == 0 {
                    0.
                } else {
                    acc.max
                }
            }
            "AVERAGE" => {
                if acc.count == 0 {
                    return Err(KernelError::new("#DIV/0!", "AVERAGE has no numeric values"));
                }
                acc.sum / acc.count as f64
            }
            "VAR" | "VARP" | "STDEV" | "STDEVP" => {
                let population = operation.ends_with('P');
                let divisor = acc.count as i64 - if population { 0 } else { 1 };
                if divisor <= 0 {
                    return Err(KernelError::new(
                        "#DIV/0!",
                        "Insufficient numeric observations",
                    ));
                }
                let variance = acc.m2 / divisor as f64;
                if operation.starts_with("STDEV") {
                    variance.max(0.).sqrt()
                } else {
                    variance
                }
            }
            _ => {
                return Err(KernelError::new(
                    "UNSUPPORTED_FEATURE",
                    format!("Missing aggregate implementation {operation}"),
                ));
            }
        };
        Ok(Value::scalar(number(value)))
    }
    fn criteria_aggregate(&self, name: &str, args: &[Value]) -> KernelResult<Value> {
        let single = matches!(name, "SUMIF" | "COUNTIF" | "AVERAGEIF");
        let count = name.starts_with("COUNT");
        let (target, pairs): (&Value, Vec<(&Value, Scalar)>) = if single {
            if args.len() < 2 || args.len() > if count { 2 } else { 3 } {
                return Err(KernelError::new(
                    "#VALUE!",
                    "Invalid criteria argument count",
                ));
            }
            let target = if args.len() == 3 { &args[2] } else { &args[0] };
            (target, vec![(&args[0], self.scalar(&args[1])?)])
        } else {
            let start = if count { 0 } else { 1 };
            if args.len() < start + 2 || (args.len() - start) % 2 != 0 {
                return Err(KernelError::new(
                    "#VALUE!",
                    "Criteria require range/value pairs",
                ));
            }
            let mut pairs = Vec::new();
            for pair in args[start..].chunks_exact(2) {
                pairs.push((&pair[0], self.scalar(&pair[1])?));
            }
            (&args[0], pairs)
        };
        let shape = self.shape(pairs[0].0);
        if pairs.iter().any(|(range, _)| self.shape(range) != shape)
            || (!count && self.shape(target) != shape)
        {
            return Err(KernelError::new(
                "#VALUE!",
                "Criteria and result ranges must have identical shapes",
            ));
        }
        let mut acc = Accumulator::default();
        let mut matched = 0u64;
        for row in 0..shape.0 {
            for col in 0..shape.1 {
                let mut pass = true;
                for (range, criterion) in &pairs {
                    if !matches_criteria(&self.at(range, row, col)?, criterion)? {
                        pass = false;
                        break;
                    }
                }
                if pass {
                    matched += 1;
                    if !count {
                        match self.at(target, row, col)? {
                            Scalar::Number(n) => acc.push(n),
                            Scalar::Error(e) => return Err(KernelError::new(e.code, e.message)),
                            _ => {}
                        }
                    }
                }
            }
        }
        let result = if count {
            matched as f64
        } else if name.starts_with("AVERAGE") {
            if acc.count == 0 {
                return Err(KernelError::new(
                    "#DIV/0!",
                    "No numeric values satisfy criteria",
                ));
            }
            acc.sum / acc.count as f64
        } else if name == "MAXIFS" {
            if acc.count == 0 { 0. } else { acc.max }
        } else if name == "MINIFS" {
            if acc.count == 0 { 0. } else { acc.min }
        } else {
            acc.sum
        };
        Ok(Value::scalar(number(result)))
    }
}
struct Accumulator {
    count: u64,
    counta: u64,
    sum: f64,
    compensation: f64,
    product: f64,
    min: f64,
    max: f64,
    mean: f64,
    m2: f64,
}
impl Default for Accumulator {
    fn default() -> Self {
        Self {
            count: 0,
            counta: 0,
            sum: 0.,
            compensation: 0.,
            product: 1.,
            min: f64::INFINITY,
            max: f64::NEG_INFINITY,
            mean: 0.,
            m2: 0.,
        }
    }
}
impl Accumulator {
    fn push(&mut self, n: f64) {
        self.count += 1;
        let corrected = n - self.compensation;
        let sum = self.sum + corrected;
        self.compensation = (sum - self.sum) - corrected;
        self.sum = sum;
        self.product *= n;
        self.min = self.min.min(n);
        self.max = self.max.max(n);
        let delta = n - self.mean;
        self.mean += delta / self.count as f64;
        self.m2 += delta * (n - self.mean);
    }
}
fn matches_criteria(value: &Scalar, criterion: &Scalar) -> KernelResult<bool> {
    if let Scalar::Error(e) = criterion {
        return Err(KernelError::new(&e.code, &e.message));
    }
    let Scalar::Text(source) = criterion else {
        return Ok(compare(value, criterion).is_eq());
    };
    let (op, operand) =
        if source.starts_with("<>") || source.starts_with("<=") || source.starts_with(">=") {
            (&source[..2], &source[2..])
        } else if source.starts_with(['=', '<', '>']) {
            (&source[..1], &source[1..])
        } else {
            ("=", source.as_str())
        };
    let target = if let Ok(n) = operand.parse::<f64>() {
        Scalar::Number(n)
    } else if operand.eq_ignore_ascii_case("TRUE") {
        Scalar::Boolean(true)
    } else if operand.eq_ignore_ascii_case("FALSE") {
        Scalar::Boolean(false)
    } else {
        Scalar::Text(operand.into())
    };
    if matches!(op, "=" | "<>") && matches!(target, Scalar::Text(_)) {
        let matched = match value {
            Scalar::Text(t) => wildcard(operand, t),
            Scalar::Null => operand.is_empty(),
            _ => false,
        };
        return Ok(if op == "<>" { !matched } else { matched });
    }
    let c = compare(value, &target);
    Ok(match op {
        "=" => c.is_eq(),
        "<>" => !c.is_eq(),
        "<" => c.is_lt(),
        ">" => c.is_gt(),
        "<=" => !c.is_gt(),
        _ => !c.is_lt(),
    })
}
fn wildcard(pattern: &str, value: &str) -> bool {
    let p: Vec<char> = pattern.to_lowercase().chars().collect();
    let v: Vec<char> = value.to_lowercase().chars().collect();
    let (mut i, mut j, mut star, mut retry) = (0, 0, None, 0);
    while j < v.len() {
        if i < p.len() && p[i] == '*' {
            star = Some(i);
            i += 1;
            retry = j;
        } else if i < p.len() && p[i] == '~' && i + 1 < p.len() && p[i + 1] == v[j] {
            i += 2;
            j += 1;
        } else if i < p.len() && (p[i] == '?' || p[i] == v[j]) {
            i += 1;
            j += 1;
        } else if let Some(s) = star {
            retry += 1;
            j = retry;
            i = s + 1;
        } else {
            return false;
        }
    }
    while i < p.len() && p[i] == '*' {
        i += 1;
    }
    i == p.len()
}
fn arity(args: &[Expr], min: usize, max: usize) -> KernelResult<()> {
    if args.len() < min || args.len() > max {
        Err(KernelError::new(
            "#VALUE!",
            format!("Expected {min} to {max} arguments"),
        ))
    } else {
        Ok(())
    }
}
fn positive_dimension(n: f64) -> KernelResult<usize> {
    if !n.is_finite() || n < 1. || n > u32::MAX as f64 {
        Err(KernelError::new(
            "#VALUE!",
            "Array dimension must be positive",
        ))
    } else {
        Ok(n.trunc() as usize)
    }
}
fn column_name(mut column: u32) -> String {
    let mut out = String::new();
    while column > 0 {
        column -= 1;
        out.insert(0, (b'A' + (column % 26) as u8) as char);
        column /= 26;
    }
    out
}
fn number(n: f64) -> Scalar {
    if n.is_finite() {
        Scalar::Number(n)
    } else {
        Scalar::error("#NUM!", "Non-finite numeric result")
    }
}
fn compare(a: &Scalar, b: &Scalar) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    match (a, b) {
        (Scalar::Null, Scalar::Number(n)) => 0f64.partial_cmp(n).unwrap_or(Ordering::Equal),
        (Scalar::Number(n), Scalar::Null) => n.partial_cmp(&0.).unwrap_or(Ordering::Equal),
        (Scalar::Null, Scalar::Text(t)) => "".cmp(t),
        (Scalar::Text(t), Scalar::Null) => t.as_str().cmp(""),
        (Scalar::Number(a), Scalar::Number(b)) => a.partial_cmp(b).unwrap_or(Ordering::Equal),
        (Scalar::Boolean(a), Scalar::Boolean(b)) => a.cmp(b),
        (Scalar::Text(a), Scalar::Text(b)) => a.to_lowercase().cmp(&b.to_lowercase()),
        (Scalar::Null, Scalar::Null) => Ordering::Equal,
        _ => {
            fn rank(v: &Scalar) -> u8 {
                match v {
                    Scalar::Null | Scalar::Number(_) => 0,
                    Scalar::Text(_) => 1,
                    Scalar::Boolean(_) => 2,
                    Scalar::Error(_) => 3,
                }
            }
            rank(a).cmp(&rank(b))
        }
    }
}
fn binary(op: &str, a: &Scalar, b: &Scalar) -> Scalar {
    if let Scalar::Error(_) = a {
        return a.clone();
    }
    if let Scalar::Error(_) = b {
        return b.clone();
    }
    if op == "&" {
        return Scalar::Text(format!("{}{}", text(a), text(b)));
    }
    if matches!(op, "=" | "<>" | "<" | ">" | "<=" | ">=") {
        let c = compare(a, b);
        return Scalar::Boolean(match op {
            "=" => c.is_eq(),
            "<>" => !c.is_eq(),
            "<" => c.is_lt(),
            ">" => c.is_gt(),
            "<=" => !c.is_gt(),
            _ => !c.is_lt(),
        });
    }
    let result = (|| {
        let a = num(a)?;
        let b = num(b)?;
        Ok(number(match op {
            "+" => a + b,
            "-" => a - b,
            "*" => a * b,
            "/" if b == 0. => return Err(KernelError::new("#DIV/0!", "Division by zero")),
            "/" => a / b,
            "^" => a.powf(b),
            _ => return Err(KernelError::new("#VALUE!", "Unknown operator")),
        }))
    })();
    match result {
        Ok(v) => v,
        Err(e) => Scalar::error(&e.code, e.message),
    }
}
