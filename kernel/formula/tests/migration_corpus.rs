use kernel_core::{Cell, CellAddress, CellReader, KernelResult, RangeRef, Scalar};
use kernel_formula::{
    CalculationServices, DefinedNameScope, FormulaRuntime, FormulaTable, FormulaValue,
    FormulaVisibility,
};
#[derive(Default)]
struct EmptyReader;

impl CellReader for EmptyReader {
    fn revision(&self) -> u64 {
        1
    }

    fn read_cell(&self, _: &CellAddress) -> KernelResult<Option<Cell>> {
        Ok(None)
    }

    fn read_range(
        &self,
        _: &RangeRef,
        _: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        Ok(())
    }
}

fn at(sheet: &str, row: u32, column: u32) -> CellAddress {
    CellAddress {
        sheet_id: sheet.into(),
        row,
        column,
    }
}

fn number(value: &FormulaValue) -> f64 {
    match value.scalar() {
        Scalar::Number(value) => value,
        other => panic!("expected number, got {other:?}"),
    }
}

fn text(value: &FormulaValue) -> String {
    match value.scalar() {
        Scalar::Text(value) => value,
        other => panic!("expected text, got {other:?}"),
    }
}

fn error(value: &FormulaValue) -> String {
    match value.scalar() {
        Scalar::Error(value) => value.code,
        other => panic!("expected error, got {other:?}"),
    }
}

#[test]
fn migration_calculation_session_is_persistent_and_delta_driven() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_value(at("S", 0, 0), Scalar::Number(2.)).unwrap();
    runtime.set_formula(at("S", 0, 1), "=A1*3").unwrap();
    let first = runtime.recalculate(&reader).unwrap();
    assert_eq!(number(&first[&at("S", 0, 1)]), 6.);
    assert_eq!(runtime.formula_count(), 1);

    let mut restored = runtime.clone();
    restored.set_value(at("S", 0, 0), Scalar::Number(5.)).unwrap();
    assert_eq!(restored.dirty_count(), 1);
    let second = restored.recalculate(&reader).unwrap();
    assert_eq!(number(&second[&at("S", 0, 1)]), 15.);
    assert!(restored.generation() > runtime.generation());
}

#[test]
fn migration_defined_names_preserve_scope_sources_and_invalidation() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    runtime.register_sheet("Sheet2", "S2").unwrap();
    runtime.set_value(at("S", 0, 0), Scalar::Number(10.)).unwrap();
    runtime.set_value(at("S", 1, 0), Scalar::Number(20.)).unwrap();
    runtime
        .define_name(
            "BaseRange",
            "=A1:A2",
            DefinedNameScope::Workbook,
            &at("S", 0, 0),
        )
        .unwrap();
    runtime
        .define_name(
            "Rate",
            "=2",
            DefinedNameScope::Workbook,
            &at("S", 0, 0),
        )
        .unwrap();
    runtime
        .define_name(
            "Rate",
            "=3",
            DefinedNameScope::Sheet("S2".into()),
            &at("S2", 0, 0),
        )
        .unwrap();
    assert_eq!(
        number(&runtime.evaluate("=SUM(BASERANGE)", &at("S", 5, 0), &reader).unwrap()),
        30.
    );
    runtime.set_formula(at("S", 0, 1), "=rate").unwrap();
    runtime.set_formula(at("S2", 0, 1), "=RATE").unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert_eq!(number(&values[&at("S", 0, 1)]), 2.);
    assert_eq!(number(&values[&at("S2", 0, 1)]), 3.);

    runtime
        .define_name(
            "RATE",
            "=4",
            DefinedNameScope::Workbook,
            &at("S", 0, 0),
        )
        .unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert_eq!(number(&values[&at("S", 0, 1)]), 4.);
    assert_eq!(number(runtime.cached_value(&at("S2", 0, 1)).unwrap()), 3.);
    assert_eq!(
        number(&runtime.clone().evaluate("=Rate", &at("S2", 1, 0), &reader).unwrap()),
        3.
    );

    runtime.set_formula(at("S", 2, 0), "=RAND()").unwrap();
    let first = runtime.recalculate(&reader).unwrap()[&at("S", 2, 0)].clone();
    runtime.set_value(at("S", 3, 0), Scalar::Number(1.)).unwrap();
    let second = runtime.recalculate(&reader).unwrap()[&at("S", 2, 0)].clone();
    assert_ne!(first, second);
}

#[test]
fn migration_external_formula_namespaces_normalize_without_touching_text() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    for (row, value) in [1., 2., 3.].into_iter().enumerate() {
        runtime
            .set_value(at("S", row as u32, 0), Scalar::Number(value))
            .unwrap();
        runtime
            .set_value(at("S", row as u32, 1), Scalar::Boolean(value != 2.))
            .unwrap();
    }
    let filtered = runtime
        .evaluate(
            "=_xlfn.FILTER(_xlws.A1:A3,B1:B3=TRUE)",
            &at("S", 0, 3),
            &reader,
        )
        .unwrap();
    assert_eq!(filtered.matrix(), vec![vec![Scalar::Number(1.)], vec![Scalar::Number(3.)]]);
    assert_eq!(
        text(
            &runtime
                .evaluate(r#"="_xlfn.FILTER"&_xlfn.SINGLE(A1)"#, &at("S", 0, 3), &reader)
                .unwrap()
        ),
        "_xlfn.FILTER1"
    );
    assert_eq!(
        number(&runtime.evaluate("=@A1", &at("S", 0, 3), &reader).unwrap()),
        1.
    );
}

#[test]
fn migration_formula_trace_uses_the_canonical_ast_evaluator() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_value(at("S", 0, 0), Scalar::Number(2.)).unwrap();
    runtime.set_value(at("S", 1, 0), Scalar::Number(3.)).unwrap();
    runtime.set_formula(at("S", 0, 1), "=SUM(A1:A2)*2").unwrap();
    let trace = runtime.trace(&at("S", 0, 1), &reader).unwrap();
    assert_eq!(number(&trace.value), 10.);
    assert!(trace.steps.len() >= 3);
    assert_eq!(trace.steps.last().unwrap().value, trace.value);
}

#[test]
fn migration_dynamic_arrays_cover_functions_empty_results_and_budgets() {
    let reader = EmptyReader;
    let runtime = FormulaRuntime::new("S");
    assert_eq!(
        runtime.evaluate("=SEQUENCE(2,3)", &at("S", 0, 0), &reader).unwrap().matrix(),
        vec![
            vec![Scalar::Number(1.), Scalar::Number(2.), Scalar::Number(3.)],
            vec![Scalar::Number(4.), Scalar::Number(5.), Scalar::Number(6.)],
        ]
    );
    assert_eq!(
        runtime
            .evaluate("=FILTER({1;2;3},{TRUE;FALSE;TRUE})", &at("S", 0, 0), &reader)
            .unwrap()
            .shape(),
        (2, 1)
    );
    assert_eq!(
        runtime
            .evaluate("=UNIQUE({1;1;2})", &at("S", 0, 0), &reader)
            .unwrap()
            .shape(),
        (2, 1)
    );
    assert_eq!(
        runtime
            .evaluate("=SORT({2;1})", &at("S", 0, 0), &reader)
            .unwrap()
            .matrix()[0][0],
        Scalar::Number(1.)
    );
    assert_eq!(
        number(&runtime.evaluate("=XMATCH(2,{1;2;3})", &at("S", 0, 0), &reader).unwrap()),
        2.
    );
    for formula in [
        "=HSTACK({1;2},{3;4})",
        "=VSTACK({1,2},{3,4})",
        "=TAKE({1;2;3},2)",
        "=DROP({1;2;3},1)",
        r#"=SORTBY({"b";"a"},{2;1})"#,
        "=RANDARRAY(2,2,1,9,TRUE)",
    ] {
        assert!(matches!(
            runtime.evaluate(formula, &at("S", 0, 0), &reader).unwrap(),
            FormulaValue::Array(_)
        ));
    }
    let empty_fallback = runtime
        .evaluate(r#"=FILTER(1,FALSE,"none")"#, &at("S", 0, 0), &reader)
        .unwrap();
    assert_eq!(
        empty_fallback.matrix(),
        vec![vec![Scalar::Text("none".into())]]
    );
    assert_eq!(
        error(
            &runtime
                .evaluate("=FILTER(1,FALSE)", &at("S", 0, 0), &reader)
                .unwrap()
        ),
        "#CALC!"
    );
    let mut bounded = FormulaRuntime::new("S");
    bounded.context.max_array_cells = 6;
    assert_eq!(
        bounded.evaluate("=SEQUENCE(2,3)", &at("S", 0, 0), &reader).unwrap().shape(),
        (2, 3)
    );
    assert_eq!(
        bounded
            .evaluate("=SEQUENCE(7)", &at("S", 0, 0), &reader)
            .unwrap_err()
            .code,
        "RESOURCE_BUDGET_EXCEEDED"
    );
}

#[test]
fn migration_groupby_and_pivotby_preserve_typed_shapes_and_aggregations() {
    let reader = EmptyReader;
    let runtime = FormulaRuntime::new("S");
    let grouped = runtime
        .evaluate(
            r#"=GROUPBY({"East","A";"East","A";"West","A";"East","B"},{10,1;20,2;5,3;7,4},SUM)"#,
            &at("S", 0, 0),
            &reader,
        )
        .unwrap();
    assert_eq!(grouped.shape(), (3, 4));
    assert_eq!(grouped.matrix()[0][2], Scalar::Number(30.));
    assert_eq!(
        runtime
            .evaluate(
                r#"=GROUPBY({"A";"A";"B"},{10;"";5},COUNT)"#,
                &at("S", 0, 0),
                &reader,
            )
            .unwrap()
            .matrix()[0][1],
        Scalar::Number(1.)
    );
    assert_eq!(
        runtime
            .evaluate(
                r#"=GROUPBY({"";"";1;"1"},{2;3;5;7},1)"#,
                &at("S", 0, 0),
                &reader,
            )
            .unwrap()
            .shape(),
        (3, 2)
    );
    assert_eq!(
        error(
            &runtime
                .evaluate(
                    r#"=GROUPBY({"A";"B"},{1},SUM)"#,
                    &at("S", 0, 0),
                    &reader,
                )
                .unwrap()
        ),
        "#VALUE!"
    );
    let pivot = runtime
        .evaluate(
            r#"=PIVOTBY({"East";"East";"West"},{"Q1";"Q2";"Q1"},{100;200;50},SUM)"#,
            &at("S", 0, 0),
            &reader,
        )
        .unwrap();
    assert_eq!(pivot.shape(), (3, 3));
    let multi = runtime
        .evaluate(
            r#"=PIVOTBY({"East",2025;"East",2025;"West",2025},{"Q1",TRUE;"Q2",TRUE;"Q1",TRUE},{10,1;20,2;5,3},SUM)"#,
            &at("S", 0, 0),
            &reader,
        )
        .unwrap();
    assert_eq!(multi.shape(), (5, 6));
    assert_eq!(
        error(
            &runtime
                .evaluate(
                    r#"=PIVOTBY({"East"},{"Q1"},{1},SUM,3)"#,
                    &at("S", 0, 0),
                    &reader,
                )
                .unwrap()
        ),
        "#VALUE!"
    );
}

#[test]
fn migration_sjs_table_is_an_isolated_what_if_array() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_formula(at("S", 0, 2), "=B1*2").unwrap();
    runtime.set_value(at("S", 1, 0), Scalar::Number(3.)).unwrap();
    runtime.set_value(at("S", 2, 0), Scalar::Number(4.)).unwrap();
    runtime
        .set_formula(at("S", 0, 3), "=SJS.TABLE(C1,A2:A3,B1)")
        .unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert_eq!(
        values[&at("S", 0, 3)].matrix(),
        vec![vec![Scalar::Number(6.)], vec![Scalar::Number(8.)]]
    );
    assert_eq!(runtime.cached_value(&at("S", 1, 3)), None);
    runtime
        .set_formula(at("S", 0, 3), "=SJS.TABLE(C1,A2:A3)")
        .unwrap();
    assert_eq!(
        error(&runtime.recalculate(&reader).unwrap()[&at("S", 0, 3)]),
        "#VALUE!"
    );
}

#[test]
fn migration_formula_index_semantics_are_native_and_fail_closed() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    runtime.register_sheet("Other", "S2").unwrap();
    runtime.set_value(at("S", 0, 0), Scalar::Number(2.)).unwrap();
    runtime.set_formula(at("S", 0, 1), "=A1+3*2").unwrap();
    runtime.set_formula(at("S", 0, 2), "=IF(FALSE,1/0,B1)").unwrap();
    runtime.set_formula(at("S", 0, 3), "=SEQUENCE(2)").unwrap();
    runtime.set_formula(at("S", 0, 4), "=SUM(D1#)").unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert_eq!(number(&values[&at("S", 0, 1)]), 8.);
    assert_eq!(number(&values[&at("S", 0, 2)]), 8.);
    assert_eq!(number(&values[&at("S", 0, 4)]), 3.);
    assert!(runtime.dependencies(&at("S", 0, 1)).unwrap().contains(&at("S", 0, 0)));
    assert!(kernel_formula::function_capabilities().iter().any(|item| item.id == "GROUPBY"));
    runtime.set_formula(at("S", 4, 0), "=A5").unwrap();
    assert_eq!(
        error(&runtime.recalculate(&reader).unwrap()[&at("S", 4, 0)]),
        "#REF!"
    );
    assert_eq!(
        runtime
            .evaluate("='Missing'!A1", &at("S", 0, 0), &reader)
            .unwrap_err()
            .code,
        "#REF!"
    );
}

#[test]
fn migration_recalculation_modes_keep_pending_and_partial_closures() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_value(at("S", 0, 0), Scalar::Number(1.)).unwrap();
    runtime.set_formula(at("S", 0, 1), "=A1+1").unwrap();
    runtime.set_formula(at("S", 0, 2), "=B1+1").unwrap();
    runtime.recalculate(&reader).unwrap();
    runtime.set_value(at("S", 0, 0), Scalar::Number(9.)).unwrap();
    assert_eq!(runtime.dirty_count(), 2);
    assert_eq!(number(runtime.cached_value(&at("S", 0, 1)).unwrap()), 2.);
    assert_eq!(number(runtime.cached_value(&at("S", 0, 2)).unwrap()), 3.);
    assert_eq!(
        number(
            &runtime
                .recalculate_cell(&at("S", 0, 1), &reader)
                .unwrap()
                .unwrap()
        ),
        10.
    );
    assert_eq!(runtime.dirty_count(), 1);
    assert_eq!(number(runtime.cached_value(&at("S", 0, 2)).unwrap()), 3.);
}

struct Visible;

impl CalculationServices for Visible {
    fn visibility(&self, _: &CellAddress) -> KernelResult<FormulaVisibility> {
        Ok(FormulaVisibility::default())
    }
}

#[test]
fn migration_structured_tables_resolve_columns_rows_and_specifiers() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    runtime
        .define_table(FormulaTable {
            name: "Sales".into(),
            range: RangeRef {
                sheet_id: "S".into(),
                start_row: 0,
                end_row: 3,
                start_column: 0,
                end_column: 1,
            },
            has_header_row: true,
            has_total_row: false,
            columns: vec!["Product".into(), "Amount".into()],
        })
        .unwrap();
    for (address, value) in [
        (at("S", 0, 0), Scalar::Text("Product".into())),
        (at("S", 0, 1), Scalar::Text("Amount".into())),
        (at("S", 1, 0), Scalar::Text("Apple".into())),
        (at("S", 1, 1), Scalar::Number(10.)),
        (at("S", 2, 0), Scalar::Text("Banana".into())),
        (at("S", 2, 1), Scalar::Number(20.)),
    ] {
        runtime.set_value(address, value).unwrap();
    }
    assert_eq!(
        number(&runtime.evaluate("=SUM(Sales[Amount])", &at("S", 0, 3), &reader).unwrap()),
        30.
    );
    assert_eq!(
        number(&runtime.evaluate("=Sales[@Amount]*2", &at("S", 1, 2), &reader).unwrap()),
        20.
    );
    assert_eq!(
        number(
            &runtime
                .evaluate_with_services(
                    "=SUBTOTAL(109,Sales[Amount])",
                    &at("S", 0, 3),
                    &reader,
                    &Visible,
                )
                .unwrap()
        ),
        30.
    );
    assert_eq!(
        text(
            &runtime
                .evaluate("=Sales[[#Headers],[Amount]]", &at("S", 0, 3), &reader)
                .unwrap()
        ),
        "Amount"
    );
    assert_eq!(
        number(&runtime.evaluate("=SUM(Sales[#Data])", &at("S", 0, 3), &reader).unwrap()),
        30.
    );
    assert_eq!(
        number(&runtime.evaluate("=SUM(Sales[#All])", &at("S", 0, 3), &reader).unwrap()),
        30.
    );
}

#[test]
fn migration_spills_share_one_owner_and_reject_blockers() {
    let reader = EmptyReader;
    let mut runtime = FormulaRuntime::new("S");
    runtime.set_formula(at("S", 0, 0), "=SEQUENCE(3)").unwrap();
    let values = runtime.recalculate(&reader).unwrap();
    assert_eq!(values[&at("S", 0, 0)].shape(), (3, 1));
    assert_eq!(
        runtime.spill_value(&at("S", 1, 0), &reader).unwrap().value,
        Some(FormulaValue::Scalar(Scalar::Number(2.)))
    );
    assert!(runtime.spill_value(&at("S", 0, 0), &reader).unwrap().is_spill);
    runtime.set_value(at("S", 1, 0), Scalar::Number(99.)).unwrap();
    runtime.set_formula(at("S", 0, 0), "=SEQUENCE(3)").unwrap();
    assert_eq!(
        error(&runtime.recalculate(&reader).unwrap()[&at("S", 0, 0)]),
        "#SPILL!"
    );
}
