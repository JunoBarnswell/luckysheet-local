use kernel_core::{Cell, CellAddress, CellReader, KernelResult, RangeRef, Scalar};
use serde_json::{Value, json};
use std::{
    cell::Cell as Counter,
    collections::BTreeMap,
    sync::atomic::{AtomicBool, Ordering},
};

struct Fixture {
    values: Vec<Vec<Scalar>>,
    scans: Counter<usize>,
}
impl CellReader for Fixture {
    fn revision(&self) -> u64 {
        7
    }
    fn read_cell(&self, address: &CellAddress) -> KernelResult<Option<Cell>> {
        Ok(self
            .values
            .get(address.row as usize)
            .and_then(|row| row.get(address.column as usize))
            .map(|value| Cell {
                value: value.clone(),
                formula: None,
                metadata: BTreeMap::new(),
            }))
    }
    fn read_range(
        &self,
        range: &RangeRef,
        visit: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        self.scans.set(self.scans.get() + 1);
        for row in range.start_row..=range.end_row {
            for column in range.start_column..=range.end_column {
                let address = CellAddress {
                    sheet_id: range.sheet_id.clone(),
                    row,
                    column,
                };
                if let Some(cell) = self.read_cell(&address)? {
                    visit(address, cell)?;
                }
            }
        }
        Ok(())
    }
}
fn fixture() -> Fixture {
    Fixture {
        values: vec![
            vec![
                Scalar::Text("A".into()),
                Scalar::Text("X".into()),
                Scalar::Number(2.),
            ],
            vec![
                Scalar::Text("A".into()),
                Scalar::Text("X".into()),
                Scalar::Text("99".into()),
            ],
            vec![
                Scalar::Text("B".into()),
                Scalar::Text("Y".into()),
                Scalar::Number(6.),
            ],
            vec![
                Scalar::Text("A".into()),
                Scalar::Text("Y".into()),
                Scalar::Number(4.),
            ],
        ],
        scans: Counter::new(0),
    }
}
fn pivot() -> Value {
    json!({"kind":"pivot","revision":7,"source":{"sheetId":"s","startRow":0,"endRow":3,"startColumn":0,"endColumn":2},"rowFields":[{"column":0}],"columnFields":[{"column":1}],"valueFields":[{"valueId":"sum","column":2,"aggregate":"sum"},{"valueId":"count","column":2,"aggregate":"count"},{"valueId":"avg","column":2,"aggregate":"average"}],"includeRowTotals":true,"includeColumnTotals":true})
}
fn cell(result: &Value, row: &[Value], column: &[Value]) -> Value {
    let row_id = result["rows"]
        .as_array()
        .unwrap()
        .iter()
        .find(|node| node["keys"] == json!(row))
        .unwrap()["rowId"]
        .as_u64()
        .unwrap();
    let column_id = result["columns"]
        .as_array()
        .unwrap()
        .iter()
        .find(|node| node["keys"] == json!(column))
        .unwrap()["columnId"]
        .as_u64()
        .unwrap();
    result["cells"]
        .as_array()
        .unwrap()
        .iter()
        .find(|cell| cell["rowId"] == row_id && cell["columnId"] == column_id)
        .unwrap()["values"]
        .clone()
}
#[test]
fn sparse_pivot_keeps_typed_values_and_repeated_placements() {
    let result = crate::execute(pivot(), &fixture()).unwrap();
    assert_eq!(
        cell(&result, &[json!("A")], &[json!("X")]),
        json!([2., 2., 2.])
    );
    assert_eq!(cell(&result, &[], &[]), json!([12., 4., 4.]));
    assert_eq!(result["cells"].as_array().unwrap().len(), 8);
}
#[test]
fn percentage_is_fraction_and_missing_base_stays_blank() {
    let mut request = pivot();
    request["valueFields"][0]["showAs"] = json!({"kind":"grand-percentage"});
    request["valueFields"][1]["showAs"] =
        json!({"kind":"difference","baseColumn":0,"baseItem":"previous"});
    let result = crate::execute(request, &fixture()).unwrap();
    assert_eq!(
        cell(&result, &[json!("A")], &[json!("X")])[0],
        json!(1. / 6.)
    );
    assert_eq!(cell(&result, &[json!("A")], &[json!("X")])[1], Value::Null);
}
#[test]
fn totals_reaggregate_after_top_filter_and_viewport_is_bounded() {
    let mut request = pivot();
    request["valueFilters"] = json!([{"axis":"rows","depth":0,"valueId":"count","top":{"direction":"top","mode":"items","threshold":1}}]);
    let result = crate::execute(request.clone(), &fixture()).unwrap();
    assert_eq!(cell(&result, &[], &[]), json!([6., 3., 3.]));
    request["viewport"] = json!({"rowOffset":0,"columnOffset":0,"rowLimit":1,"columnLimit":1});
    let page = crate::execute(request, &fixture()).unwrap();
    assert_eq!(page["rows"].as_array().unwrap().len(), 1);
    assert_eq!(page["columns"].as_array().unwrap().len(), 1);
    assert_eq!(page["cells"].as_array().unwrap().len(), 1);
}
#[test]
fn duplicate_value_ids_and_out_of_range_reject_before_publication() {
    let source = fixture();
    let mut request = pivot();
    request["valueFields"][1]["valueId"] = json!("sum");
    assert_eq!(
        crate::execute(request, &source).unwrap_err().code,
        "PIVOT_VALUE_ID_INVALID"
    );
    assert_eq!(source.scans.get(), 0);
    let mut request = pivot();
    request["rowFields"][0]["column"] = json!(10);
    assert_eq!(
        crate::execute(request, &source).unwrap_err().code,
        "ANALYTICS_COLUMN_OUTSIDE_RANGE"
    );
}
#[test]
fn source_cache_survives_layout_change_and_cancel() {
    let source = fixture();
    let runtime = crate::AnalyticsRuntime::new();
    let cancel = AtomicBool::new(false);
    runtime.execute(pivot(), &source, &cancel).unwrap();
    let mut changed = pivot();
    changed["valueFields"][0]["aggregate"] = json!("max");
    runtime.execute(changed.clone(), &source, &cancel).unwrap();
    assert_eq!(source.scans.get(), 1);
    cancel.store(true, Ordering::Release);
    assert_eq!(
        runtime
            .execute(changed.clone(), &source, &cancel)
            .unwrap_err()
            .code,
        "ANALYTICS_TASK_CANCELLED"
    );
    cancel.store(false, Ordering::Release);
    runtime.execute(changed, &source, &cancel).unwrap();
    assert_eq!(source.scans.get(), 1);
}
#[test]
fn filter_domain_excludes_own_predicate_and_respects_other_columns() {
    let source = fixture();
    let result=crate::execute(json!({"kind":"filter","revision":7,"range":{"sheetId":"s","startRow":0,"endRow":3,"startColumn":0,"endColumn":2},"conditions":[{"column":0,"predicate":{"kind":"values","values":["A"],"includeBlank":false}},{"column":1,"predicate":{"kind":"values","values":["Y"],"includeBlank":false}}]}),&source).unwrap();
    assert_eq!(result["visibleRows"], json!([3]));
    let domain = result["domain"]["0"].as_array().unwrap();
    assert!(domain.iter().any(|entry| entry["value"] == "B"));
    assert_eq!(domain.len(), 2);
}
#[test]
fn numeric_custom_filter_is_not_silently_false() {
    let result=crate::execute(json!({"kind":"filter","revision":7,"range":{"sheetId":"s","startRow":0,"endRow":3,"startColumn":0,"endColumn":2},"conditions":[{"column":2,"predicate":{"kind":"custom","join":"and","conditions":[{"operator":"greaterThan","value":3}]}}]}),&fixture()).unwrap();
    assert_eq!(result["visibleRows"], json!([2, 3]));
}

struct Generated {
    rows: u32,
    scanned: Counter<u32>,
}

struct Interrupting<'a> {
    cancel: &'a AtomicBool,
    visited: Counter<usize>,
}
impl CellReader for Interrupting<'_> {
    fn revision(&self) -> u64 {
        7
    }
    fn read_cell(&self, _: &CellAddress) -> KernelResult<Option<Cell>> {
        unreachable!()
    }
    fn read_range(
        &self,
        range: &RangeRef,
        visit: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        for row in range.start_row..=range.end_row {
            for column in range.start_column..=range.end_column {
                self.visited.set(self.visited.get() + 1);
                if self.visited.get() == 512 {
                    self.cancel.store(true, Ordering::Release);
                }
                visit(
                    CellAddress {
                        sheet_id: range.sheet_id.clone(),
                        row,
                        column,
                    },
                    Cell {
                        value: Scalar::Number(1.),
                        formula: None,
                        metadata: BTreeMap::new(),
                    },
                )?;
            }
        }
        Ok(())
    }
}
#[test]
fn cancellation_interrupts_a_large_scan_at_bounded_checkpoint() {
    let cancel = AtomicBool::new(false);
    let source = Interrupting {
        cancel: &cancel,
        visited: Counter::new(0),
    };
    let mut request = pivot();
    request["source"]["endRow"] = json!(999_999);
    let error = crate::execute_with_cancel(request, &source, &cancel).unwrap_err();
    assert_eq!(error.code, "ANALYTICS_TASK_CANCELLED");
    assert_eq!(source.visited.get(), 512);
}
#[test]
fn memory_rejection_occurs_before_source_allocation() {
    let source = fixture();
    let mut request = pivot();
    request["budget"] = json!({"memoryBytes":1,"timeoutMs":1000,"temporaryBytes":0});
    assert_eq!(
        crate::execute(request, &source).unwrap_err().code,
        "ANALYTICS_MEMORY_BUDGET_EXCEEDED"
    );
    assert_eq!(source.scans.get(), 0);
}
impl CellReader for Generated {
    fn revision(&self) -> u64 {
        7
    }
    fn read_cell(&self, _: &CellAddress) -> KernelResult<Option<Cell>> {
        unreachable!("analytics source uses range scan")
    }
    fn read_range(
        &self,
        range: &RangeRef,
        visit: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        for row in 0..self.rows {
            for column in range.start_column..=range.end_column {
                self.scanned.set(self.scanned.get() + 1);
                let value = match column {
                    0 => Scalar::Number((row % 1000) as f64),
                    1 => Scalar::Number((row % 10) as f64),
                    _ => Scalar::Number(1.),
                };
                visit(
                    CellAddress {
                        sheet_id: range.sheet_id.clone(),
                        row,
                        column,
                    },
                    Cell {
                        value,
                        formula: None,
                        metadata: BTreeMap::new(),
                    },
                )?;
            }
        }
        Ok(())
    }
}
#[test]
fn million_rows_remain_sparse_and_drilldown_is_paginated() {
    let source = Generated {
        rows: 1_000_000,
        scanned: Counter::new(0),
    };
    let mut request = pivot();
    request["source"]["endRow"] = json!(999_999);
    request["includeRowTotals"] = json!(false);
    request["includeColumnTotals"] = json!(false);
    request["viewport"] = json!({"rowOffset":0,"columnOffset":0,"rowLimit":10,"columnLimit":10});
    request["drilldown"] = json!({"rowKeys":[0],"columnKeys":[0],"offset":500,"limit":3});
    let result = crate::execute(request, &source).unwrap();
    assert_eq!(result["totalGroups"], 1000);
    assert_eq!(result["totalColumns"], 10);
    assert_eq!(result["cells"].as_array().unwrap().len(), 10);
    assert_eq!(result["drilldown"]["total"], 1000);
    assert_eq!(
        result["drilldown"]["sourceRows"],
        json!([500_000, 501_000, 502_000])
    );
    assert_eq!(source.scanned.get(), 3_000_000);
}

#[test]
fn query_hash_join_preserves_multiplicity_and_group_aggregates_numbers_only() {
    let range = json!({"sheetId":"s","startRow":0,"endRow":3,"startColumn":0,"endColumn":2});
    let mut request = json!({"kind":"query","revision":7,"range":range,"limit":20,"joins":[{"range":range,"leftColumn":1,"rightColumn":1,"kind":"inner","outputStartColumn":3}],"columns":[0,2,5]});
    let source = fixture();
    let result = crate::execute(request.clone(), &source).unwrap();
    assert_eq!(result["total"], 8);
    assert_eq!(result["rows"][0]["sourceRowIds"], json!([0, 0]));
    request["groupBy"] = json!([0]);
    request["aggregates"] = json!([{"column":5,"aggregate":"sum","outputColumn":6}]);
    request["columns"] = json!([0, 6]);
    let grouped = crate::execute(request, &source).unwrap();
    assert_eq!(grouped["total"], 2);
    assert_eq!(grouped["rows"][0]["values"], json!(["A", 14.]));
    assert_eq!(grouped["rows"][1]["values"], json!(["B", 10.]));
}
#[test]
fn grouped_query_rejects_arbitrary_ungrouped_projection() {
    let request = json!({"kind":"query","revision":7,"range":{"sheetId":"s","startRow":0,"endRow":3,"startColumn":0,"endColumn":2},"limit":20,"groupBy":[0],"columns":[0,2]});
    assert_eq!(
        crate::execute(request, &fixture()).unwrap_err().code,
        "QUERY_GROUP_PROJECTION_INVALID"
    );
}
