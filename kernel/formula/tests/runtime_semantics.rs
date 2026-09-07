use kernel_core::{Cell, CellAddress, CellReader, KernelError, KernelResult, RangeRef, Scalar};
use kernel_formula::{CalculationServices, FormulaRuntime, FormulaValue, FormulaVisibility};
use std::collections::BTreeMap;

#[derive(Default)]
struct SparseReader {
    cells: BTreeMap<CellAddress, Cell>,
    cell_reads: std::cell::Cell<usize>,
    range_reads: std::cell::Cell<usize>,
}
impl SparseReader {
    fn put(&mut self, row: u32, col: u32, value: Scalar) {
        self.cells.insert(
            CellAddress {
                sheet_id: "S".into(),
                row,
                column: col,
            },
            Cell {
                value,
                ..Default::default()
            },
        );
    }
}
impl CellReader for SparseReader {
    fn revision(&self) -> u64 {
        1
    }
    fn read_cell(&self, address: &CellAddress) -> KernelResult<Option<Cell>> {
        self.cell_reads.set(self.cell_reads.get() + 1);
        Ok(self.cells.get(address).cloned())
    }
    fn read_range(
        &self,
        range: &RangeRef,
        visitor: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        self.range_reads.set(self.range_reads.get() + 1);
        for row in range.start_row..=range.end_row {
            for col in range.start_column..=range.end_column {
                let address = CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row,
                    column: col,
                };
                if let Some(cell) = self.cells.get(&address) {
                    visitor(address, cell.clone())?;
                }
            }
        }
        Ok(())
    }
}

fn addr(row: u32, col: u32) -> CellAddress {
    CellAddress {
        sheet_id: "S".into(),
        row,
        column: col,
    }
}
fn number(v: &FormulaValue) -> f64 {
    match v.scalar() {
        Scalar::Number(n) => n,
        other => panic!("expected number, got {other:?}"),
    }
}

#[test]
fn dirty_point_and_range_dependencies_recalculate_only_the_local_closure() {
    let reader = SparseReader::default();
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_value(addr(0, 0), Scalar::Number(1.)).unwrap();
    runtime.set_formula(addr(0, 1), "=A1+1").unwrap();
    runtime.set_formula(addr(0, 2), "=B1+1").unwrap();
    runtime.set_formula(addr(0, 3), "=SUM(A1:A10)").unwrap();
    for row in 20..1020 {
        runtime.set_formula(addr(row, 10), "=1+2").unwrap();
    }
    runtime.recalculate(&reader).unwrap();
    runtime.set_value(addr(0, 0), Scalar::Number(5.)).unwrap();
    assert_eq!(runtime.dirty_count(), 3);
    reader.cell_reads.set(0);
    let changed = runtime.recalculate(&reader).unwrap();
    assert_eq!(changed.len(), 3);
    assert_eq!(number(&changed[&addr(0, 2)]), 7.);
    assert_eq!(number(&changed[&addr(0, 3)]), 5.);
    assert!(reader.cell_reads.get() < 20);
}

#[test]
fn spill_children_references_sums_resize_and_blockers_share_one_owner() {
    let reader = SparseReader::default();
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_formula(addr(0, 0), "=SEQUENCE(3)").unwrap();
    runtime.set_formula(addr(0, 2), "=A3*2").unwrap();
    runtime.set_formula(addr(0, 3), "=SUM(A1:A3)").unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert_eq!(number(&values[&addr(0, 2)]), 6.);
    assert_eq!(number(&values[&addr(0, 3)]), 6.);
    assert_eq!(
        number(&runtime.evaluate("=A2", &addr(5, 5), &reader).unwrap()),
        2.
    );
    assert!(matches!(
        runtime.evaluate("=A1", &addr(5, 5), &reader).unwrap(),
        FormulaValue::Scalar(_)
    ));
    runtime.set_formula(addr(0, 0), "=SEQUENCE(2)").unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert_eq!(number(&values[&addr(0, 2)]), 0.);
    assert_eq!(number(&values[&addr(0, 3)]), 3.);
    runtime.set_value(addr(1, 0), Scalar::Number(99.)).unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert!(matches!(values[&addr(0,0)].scalar(),Scalar::Error(e) if e.code=="#SPILL!"));
    runtime.set_value(addr(1, 0), Scalar::Null).unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert!(matches!(values[&addr(0, 0)], FormulaValue::Array(_)));
}

#[test]
fn lookup_binary_both_directions_and_matched_error_do_not_use_not_found() {
    let reader = SparseReader::default();
    let runtime = FormulaRuntime::new("S");
    let at = addr(8, 8);
    for (formula, expected) in [
        ("=XLOOKUP(20,{10;20;30},{1;2;3},0,0,2)", 2.),
        ("=XLOOKUP(20,{30;20;10},{3;2;1},0,0,-2)", 2.),
        ("=XLOOKUP(15,{10;20;30},{1;2;3},0,-1,2)", 1.),
        ("=XLOOKUP(15,{30;20;10},{3;2;1},0,1,-2)", 2.),
    ] {
        assert_eq!(
            number(&runtime.evaluate(formula, &at, &reader).unwrap()),
            expected,
            "{formula}"
        );
    }
    let result = runtime
        .evaluate("=XLOOKUP(1,{1;2},{#N/A;2},99)", &at, &reader)
        .unwrap();
    assert!(matches!(result.scalar(),Scalar::Error(e) if e.code=="#N/A"));
    assert_eq!(
        number(
            &runtime
                .evaluate("=XLOOKUP(1,{1;2},{4;5},1/0)", &at, &reader)
                .unwrap()
        ),
        4.
    );
    assert!(
        matches!(runtime.evaluate("=XLOOKUP(1,{1,2;3,4},{1,2;3,4})",&at,&reader).unwrap().scalar(),Scalar::Error(e) if e.code=="#VALUE!")
    );
}

#[test]
fn lambda_scope_recursion_omitted_shape_and_resource_rejection() {
    let reader = SparseReader::default();
    let mut runtime = FormulaRuntime::new("S");
    let at = addr(9, 9);
    assert_eq!(
        number(
            &runtime
                .evaluate("=LET(f,LAMBDA(n,IF(n=0,1,n*f(n-1))),f(5))", &at, &reader)
                .unwrap()
        ),
        120.
    );
    assert_eq!(
        runtime
            .evaluate("=LAMBDA(x,ISOMITTED(x))()", &at, &reader)
            .unwrap()
            .is_error(),
        true
    );
    assert!(
        matches!(runtime.evaluate("=MAP({1;2},LAMBDA(x,{1,2}))",&at,&reader).unwrap().scalar(),Scalar::Error(e) if e.code=="#CALC!")
    );
    runtime.context.max_array_cells = 2;
    let error = runtime
        .evaluate("=SEQUENCE(1000000000)", &at, &reader)
        .unwrap_err();
    assert_eq!(error.code, "RESOURCE_BUDGET_EXCEEDED");
}

struct MissingPage;
impl CellReader for MissingPage {
    fn revision(&self) -> u64 {
        1
    }
    fn read_cell(&self, _: &CellAddress) -> KernelResult<Option<Cell>> {
        Err(KernelError::new(
            "DATA_PAGE_UNAVAILABLE",
            "Missing revision-pinned page",
        ))
    }
    fn read_range(
        &self,
        _: &RangeRef,
        _: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        Err(KernelError::new(
            "DATA_PAGE_UNAVAILABLE",
            "Missing revision-pinned page",
        ))
    }
}
#[test]
fn host_failures_are_not_excel_errors_and_failed_recalculation_is_atomic() {
    let mut runtime = FormulaRuntime::new("S");
    let good = SparseReader::default();
    runtime.set_formula(addr(0, 1), "=A1+1").unwrap();
    runtime.recalculate(&good).unwrap();
    let generation = runtime.generation();
    runtime.invalidate(&addr(0, 0));
    assert_eq!(
        runtime
            .evaluate("=IFERROR(A1,9)", &addr(0, 1), &MissingPage)
            .unwrap_err()
            .code,
        "DATA_PAGE_UNAVAILABLE"
    );
    assert_eq!(
        runtime.recalculate(&MissingPage).unwrap_err().code,
        "DATA_PAGE_UNAVAILABLE"
    );
    assert_eq!(runtime.generation(), generation);
    assert_eq!(runtime.dirty_count(), 1);
}

#[test]
fn what_if_overrides_invalidate_cached_transitive_dependents_without_mutation() {
    let reader = SparseReader::default();
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_value(addr(0, 0), Scalar::Number(2.)).unwrap();
    runtime.set_formula(addr(0, 1), "=A1*3").unwrap();
    runtime.recalculate(&reader).unwrap();
    let generation = runtime.generation();
    let overrides = BTreeMap::from([(addr(0, 0), Scalar::Number(7.))]);
    assert_eq!(
        number(
            &runtime
                .evaluate_with_overrides("=B1", &addr(8, 8), &reader, &overrides)
                .unwrap()
        ),
        21.
    );
    assert_eq!(
        number(&runtime.evaluate("=B1", &addr(8, 8), &reader).unwrap()),
        6.
    );
    assert_eq!(runtime.generation(), generation);
}

#[test]
fn random_context_is_independent_of_calculation_visit_order() {
    let reader = SparseReader::default();
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_formula(addr(0, 0), "=RAND()+RAND()").unwrap();
    runtime.set_formula(addr(0, 1), "=RAND()").unwrap();
    let expected = runtime.evaluate("=RAND()", &addr(0, 1), &reader).unwrap();
    let results = runtime.recalculate(&reader).unwrap();
    assert_eq!(results[&addr(0, 1)], expected);
}

struct VisibilityServices;
impl CalculationServices for VisibilityServices {
    fn visibility(&self, address: &CellAddress) -> KernelResult<FormulaVisibility> {
        Ok(FormulaVisibility {
            filter_hidden: address.row == 2,
            manual_hidden: address.row == 1,
            outline_hidden: false,
        })
    }
}
#[test]
fn visibility_services_and_spill_inspection_propagate_the_owning_contract() {
    let mut reader = SparseReader::default();
    for (row, value) in [10., 20., 40.].into_iter().enumerate() {
        reader.put(row as u32, 0, Scalar::Number(value));
    }
    let mut runtime = FormulaRuntime::new("S");
    assert_eq!(
        number(
            &runtime
                .evaluate_with_services(
                    "=SUBTOTAL(9,A1:A3)",
                    &addr(8, 8),
                    &reader,
                    &VisibilityServices
                )
                .unwrap()
        ),
        30.
    );
    assert_eq!(
        number(
            &runtime
                .evaluate_with_services(
                    "=SUBTOTAL(109,A1:A3)",
                    &addr(8, 8),
                    &reader,
                    &VisibilityServices
                )
                .unwrap()
        ),
        10.
    );
    assert_eq!(
        runtime
            .evaluate("=SUBTOTAL(9,A1:A3)", &addr(8, 8), &reader)
            .unwrap_err()
            .code,
        "FORMULA_VISIBILITY_UNAVAILABLE"
    );
    runtime.set_formula(addr(0, 3), "=SEQUENCE(2)").unwrap();
    runtime.recalculate(&reader).unwrap();
    runtime.invalidate(&addr(0, 3));
    assert_eq!(
        runtime
            .spill_value_with_services(&addr(1, 3), &MissingPage, &VisibilityServices)
            .unwrap_err()
            .code,
        "DATA_PAGE_UNAVAILABLE"
    );
}

#[test]
fn reference_union_index_and_two_dimensional_logical_arrays_are_reference_aware() {
    let mut reader = SparseReader::default();
    reader.put(0, 0, Scalar::Number(1.));
    reader.put(1, 0, Scalar::Number(2.));
    reader.put(0, 1, Scalar::Number(3.));
    reader.put(1, 1, Scalar::Number(4.));
    let runtime = FormulaRuntime::new("S");
    assert_eq!(
        number(
            &runtime
                .evaluate("=SUM((A1:A2,B1:B2))", &addr(7, 7), &reader)
                .unwrap()
        ),
        10.
    );
    assert_eq!(
        number(
            &runtime
                .evaluate("=INDEX((A1:A2,B1:B2),2,1,2)", &addr(7, 7), &reader)
                .unwrap()
        ),
        4.
    );
    assert_eq!(
        runtime
            .evaluate("=AND({TRUE,TRUE;TRUE,FALSE})", &addr(7, 7), &reader)
            .unwrap()
            .scalar(),
        Scalar::Boolean(false)
    );
}

#[test]
fn logical_range_does_not_hide_terminal_error() {
    let mut reader = SparseReader::default();
    reader.put(0, 0, Scalar::Boolean(true));
    reader.put(1, 0, Scalar::Boolean(false));
    reader.put(2, 0, Scalar::error("#N/A", "terminal"));
    let runtime = FormulaRuntime::new("S");
    let result = runtime.evaluate("=AND(A1:A3)", &addr(0, 1), &reader);
    assert!(
        matches!(result.unwrap().scalar(),Scalar::Error(e) if e.code=="#N/A"),
        "AND must evaluate the final error cell as an Excel cell error"
    );
}

#[test]
fn address_modes_two_and_three() {
    let runtime = FormulaRuntime::new("S");
    let reader = SparseReader::default();
    assert_eq!(
        runtime
            .evaluate("=ADDRESS(2,3,2)", &addr(0, 0), &reader)
            .unwrap()
            .scalar(),
        Scalar::Text("C$2".into())
    );
    assert_eq!(
        runtime
            .evaluate("=ADDRESS(2,3,3)", &addr(0, 0), &reader)
            .unwrap()
            .scalar(),
        Scalar::Text("$C2".into())
    );
}

#[test]
fn offset_inherits_source_range_size() {
    let mut reader = SparseReader::default();
    reader.put(0, 1, Scalar::Number(10.));
    reader.put(1, 1, Scalar::Number(20.));
    let runtime = FormulaRuntime::new("S");
    let value = runtime
        .evaluate("=SUM(OFFSET(B1:B2,0,0))", &addr(0, 0), &reader)
        .unwrap();
    assert_eq!(number(&value), 30.);
}

#[test]
fn xlookup_modes_search_wildcard_duplicate_and_invalid() {
    let mut reader = SparseReader::default();
    for (row, key) in ["alpha", "alpha", "beta", "delta"].iter().enumerate() {
        reader.put(row as u32, 0, Scalar::Text((*key).into()));
        reader.put(row as u32, 1, Scalar::Number((row + 1) as f64));
    }
    let runtime = FormulaRuntime::new("S");
    let at = addr(0, 4);
    assert_eq!(
        number(
            &runtime
                .evaluate("=XLOOKUP(\"alpha\",A1:A4,B1:B4)", &at, &reader)
                .unwrap()
        ),
        1.
    );
    assert_eq!(
        number(
            &runtime
                .evaluate(
                    "=XLOOKUP(\"alpha\",A1:A4,B1:B4,\"missing\",0,-1)",
                    &at,
                    &reader
                )
                .unwrap()
        ),
        2.
    );
    assert_eq!(
        number(
            &runtime
                .evaluate("=XLOOKUP(\"alp*\",A1:A4,B1:B4,\"missing\",2)", &at, &reader)
                .unwrap()
        ),
        1.
    );
    assert!(
        runtime
            .evaluate(
                "=XLOOKUP(\"alpha\",A1:A4,B1:B4,\"missing\",9)",
                &at,
                &reader
            )
            .unwrap()
            .is_error()
    );
}

#[test]
fn indirect_supports_r1c1_relative_and_absolute() {
    let mut reader = SparseReader::default();
    reader.put(1, 1, Scalar::Number(42.));
    reader.put(2, 2, Scalar::Number(77.));
    let runtime = FormulaRuntime::new("S");
    assert_eq!(
        number(
            &runtime
                .evaluate("=INDIRECT(\"R2C2\",FALSE)", &addr(0, 0), &reader)
                .unwrap()
        ),
        42.
    );
    assert_eq!(
        number(
            &runtime
                .evaluate("=INDIRECT(\"R[1]C[1]\",FALSE)", &addr(1, 1), &reader)
                .unwrap()
        ),
        77.
    );
}

#[test]
fn modern_lambda_array_functions_have_real_results() {
    let reader = SparseReader::default();
    let runtime = FormulaRuntime::new("S");
    let at = addr(0, 0);
    assert_eq!(
        number(&runtime.evaluate("=LET(x,10,x+20)", &at, &reader).unwrap()),
        30.
    );
    assert_eq!(
        runtime
            .evaluate("=MAP({1,2,3},LAMBDA(x,x*2))", &at, &reader)
            .unwrap()
            .shape(),
        (1, 3)
    );
    assert_eq!(
        runtime
            .evaluate("=REDUCE(0,{1,2,3},LAMBDA(a,b,a+b))", &at, &reader)
            .unwrap()
            .scalar(),
        Scalar::Number(6.)
    );
    assert_eq!(
        runtime
            .evaluate("=SCAN(0,{1,2,3},LAMBDA(a,b,a+b))", &at, &reader)
            .unwrap()
            .shape(),
        (1, 3)
    );
    assert_eq!(
        runtime
            .evaluate("=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))", &at, &reader)
            .unwrap()
            .shape(),
        (2, 1)
    );
    assert_eq!(
        runtime
            .evaluate("=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))", &at, &reader)
            .unwrap()
            .shape(),
        (1, 2)
    );
    assert_eq!(
        runtime
            .evaluate("=MAKEARRAY(2,2,LAMBDA(r,c,r+c))", &at, &reader)
            .unwrap()
            .shape(),
        (2, 2)
    );
}

struct Cancel;
impl CalculationServices for Cancel {
    fn checkpoint(&self, _: u64) -> KernelResult<()> {
        Err(KernelError::new("CANCELLED", "test cancellation"))
    }
    fn visibility(&self, _: &CellAddress) -> KernelResult<FormulaVisibility> {
        Ok(FormulaVisibility::default())
    }
}

#[test]
fn cancellation_does_not_update_cache() {
    let reader = SparseReader::default();
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_formula(addr(0, 0), "=SEQUENCE(2,2)").unwrap();
    assert_eq!(
        runtime
            .recalculate_with_services(&reader, &Cancel)
            .unwrap_err()
            .code,
        "CANCELLED"
    );
    assert_eq!(runtime.dirty_count(), 1);
}
