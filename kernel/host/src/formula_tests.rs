use super::*;

fn create(host: &mut KernelHost, id: &str) {
    host.dispatch(
        "create",
        json!({
            "unitId": id,
            "name": "Formula host",
            "sheets": [{
                "sheetId": "s",
                "name": "Sheet1",
                "rowCount": 20,
                "columnCount": 10,
                "metadata": {}
            }]
        }),
    )
    .unwrap();
}

fn set_cell(
    host: &mut KernelHost,
    id: &str,
    revision: u64,
    operation: &str,
    row: u32,
    column: u32,
    value: Value,
) -> KernelResult<Value> {
    host.dispatch(
        "command",
        json!({
            "unitId": id,
            "baseRevision": revision,
            "operationId": operation,
            "commandId": "cell.set",
            "accessRole": "editor",
            "params": {
                "sheetId": "s",
                "row": row,
                "column": column,
                "value": value
            }
        }),
    )
}

#[test]
fn formula_host_routes_share_one_runtime_across_command_revisions() {
    let mut host = KernelHost::default();
    create(&mut host, "formula");
    set_cell(&mut host, "formula", 0, "value-1", 0, 0, json!({"value": 2})).unwrap();
    set_cell(
        &mut host,
        "formula",
        1,
        "formula-1",
        0,
        1,
        json!({"value": null, "formula": "=A1*3"}),
    )
    .unwrap();

    let first = host
        .dispatch(
            "formula.recalculate",
            json!({"unitId":"formula","revision":2}),
        )
        .unwrap();
    assert_eq!(first["recalculatedCount"], 1);
    let inspection = host
        .dispatch(
            "formula.inspect",
            json!({
                "unitId":"formula",
                "revision":2,
                "address":{"sheetId":"s","row":0,"column":1}
            }),
        )
        .unwrap();
    assert_eq!(inspection["entries"][0]["value"].as_f64(), Some(6.));

    set_cell(&mut host, "formula", 2, "value-2", 0, 0, json!({"value": 5})).unwrap();
    let second = host
        .dispatch(
            "formula.recalculate",
            json!({"unitId":"formula","revision":3}),
        )
        .unwrap();
    assert_eq!(second["recalculatedCount"], 1);
    let trace = host
        .dispatch(
            "formula.trace",
            json!({
                "unitId":"formula",
                "revision":3,
                "address":{"sheetId":"s","row":0,"column":1}
            }),
        )
        .unwrap();
    assert_eq!(trace["value"].as_f64(), Some(15.));
    assert!(trace["steps"].as_array().is_some_and(|steps| !steps.is_empty()));
}

#[test]
fn rejected_command_leaves_formula_registry_and_revision_unchanged() {
    let mut host = KernelHost::default();
    create(&mut host, "atomic-formula");
    set_cell(
        &mut host,
        "atomic-formula",
        0,
        "formula",
        0,
        0,
        json!({"value": null, "formula": "=1+2"}),
    )
    .unwrap();
    host.dispatch(
        "formula.recalculate",
        json!({"unitId":"atomic-formula","revision":1}),
    )
    .unwrap();
    let before = host
        .dispatch(
            "formula.inspect",
            json!({"unitId":"atomic-formula","revision":1,"projection":"entries"}),
        )
        .unwrap();
    let error = set_cell(
        &mut host,
        "atomic-formula",
        0,
        "stale",
        0,
        0,
        json!({"value": 99}),
    )
    .unwrap_err();
    assert_eq!(error.code, "STALE_REVISION");
    let after = host
        .dispatch(
            "formula.inspect",
            json!({"unitId":"atomic-formula","revision":1,"projection":"entries"}),
        )
        .unwrap();
    assert_eq!(after, before);
}
