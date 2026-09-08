use super::*;

#[test]
fn native_creation_owns_default_sheet_and_rejects_invalid_intent_atomically() {
    let mut host = KernelHost::default();
    for sheets in [Value::Null, json!([]), json!([{"sheetId":"bad"}])] {
        assert!(host.dispatch("create", json!({"unitId":"new","name":"New","sheets":sheets})).is_err());
        assert!(!host.workbooks.contains_key("new"));
    }
    let manifest = host.dispatch("create", json!({"unitId":"new","name":"New"})).unwrap();
    assert_eq!(manifest["revision"], 0);
    assert_eq!(manifest["sheets"][0]["rowCount"], MAX_ROWS);
    assert_eq!(manifest["sheets"][0]["columnCount"], MAX_COLUMNS);
    assert_eq!(manifest["pages"], json!([]));
    assert_eq!(host.dispatch("create", json!({"unitId":"new","name":"Replacement"})).unwrap_err().code, "WORKBOOK_ALREADY_OPEN");
    assert_eq!(host.dispatch("manifest", json!({"unitId":"new","revision":0})).unwrap(), manifest);
}

fn open_fixture(host: &mut KernelHost, id: &str, value: f64) {
    let mut workbook = WorkbookPages::create(id, id, vec![SheetManifest {
        sheet_id: "s".into(), name: "Sheet1".into(), row_count: 10,
        column_count: 2, metadata: BTreeMap::new(),
    }]).unwrap();
    workbook.apply_writes("seed", 0, vec![CellWrite {
        address: CellAddress { sheet_id: "s".into(), row: 0, column: 0 },
        cell: Some(Cell { value: Scalar::Number(value), formula: None, metadata: BTreeMap::new() }),
    }]).unwrap();
    host.workbooks.insert(id.into(), workbook);
}

fn query(id: &str, revision: u64) -> Value {
    json!({"unitId":id,"revision":revision,"request":{
        "kind":"query","revision":revision,"range":{"sheetId":"s","startRow":0,
        "endRow":0,"startColumn":0,"endColumn":0},"limit":1
    }})
}

#[test]
fn restore_publishes_new_revision_and_rejects_unauthorized_or_foreign_history() {
    let mut host = KernelHost::default();
    open_fixture(&mut host, "restore", 5.);
    let target = host.workbooks["restore"].manifest();
    host.workbooks.get_mut("restore").unwrap().apply_writes("edit", 1, vec![CellWrite {
        address: CellAddress { sheet_id: "s".into(), row: 0, column: 0 },
        cell: Some(Cell { value: Scalar::Number(9.), formula: None, metadata: BTreeMap::new() }),
    }]).unwrap();
    let current = host.workbooks["restore"].manifest();
    let mut request = json!({"unitId":"restore","baseRevision":2,"operationId":"restore-old",
        "accessRole":"editor","targetManifest":target});
    assert_eq!(host.dispatch("restore", request.clone()).unwrap_err().code, "FORBIDDEN");
    request["accessRole"] = json!("owner");
    request["targetManifest"]["unitId"] = json!("foreign");
    assert_eq!(host.dispatch("restore", request.clone()).unwrap_err().code, "HISTORY_INVALID");
    assert_eq!(host.workbooks["restore"].manifest(), current);
    request["targetManifest"]["unitId"] = json!("restore");
    let result = host.dispatch("restore", request.clone()).unwrap();
    assert_eq!(result["revision"], 3);
    assert_eq!(result["baseRevision"], 2);
    assert_eq!(result["pages"], json!([]));
    assert_eq!(result["manifest"]["pages"], json!(target.pages));
    assert_eq!(result["history"]["pageDeltas"][0]["before"], json!(current.pages[0]));
    assert_eq!(result["history"]["pageDeltas"][0]["after"], json!(target.pages[0]));
    assert_eq!(result["history"]["requiredRole"], json!("owner"));
    assert_eq!(host.dispatch("restore", request).unwrap_err().code, "STALE_REVISION");
}

#[test]
fn copy_preserves_page_content_and_rejects_stale_or_existing_targets() {
    let mut host = KernelHost::default();
    open_fixture(&mut host, "source", 23.);
    let mut request = json!({"sourceUnitId":"source","sourceRevision":0,"targetUnitId":"copy","name":"Copy"});
    assert_eq!(host.dispatch("copy", request.clone()).unwrap_err().code, "STALE_REVISION");
    assert!(!host.workbooks.contains_key("copy"));
    request["sourceRevision"] = json!(1);
    let source_manifest = host.workbooks["source"].manifest();
    let result = host.dispatch("copy", request.clone()).unwrap();
    assert_eq!(result["manifest"]["revision"], 0);
    assert_eq!(result["manifest"]["pages"][0]["checksum"], source_manifest.pages[0].checksum);
    assert_eq!(result["manifest"]["pages"][0]["revision"], 0);
    assert_eq!(host.dispatch("analytics.execute", query("copy", 0)).unwrap_err().code, "DATA_PAGE_UNAVAILABLE");
    let descriptor = host.workbooks["copy"].manifest().pages[0].clone();
    let bytes = host.workbooks["source"].page_bytes(&descriptor.key()).unwrap();
    host.dispatch("page.load", json!({"unitId":"copy","revision":0,"page":PagePayload {
        descriptor, payload_base64: STANDARD.encode(bytes)
    }})).unwrap();
    assert_eq!(host.dispatch("analytics.execute", query("copy", 0)).unwrap()["rows"][0]["values"], json!([23.]));
    assert_eq!(host.dispatch("copy", request).unwrap_err().code, "WORKBOOK_ALREADY_OPEN");
    assert_eq!(host.workbooks["source"].manifest(), source_manifest);
}

#[test]
fn analytics_workbooks_are_isolated_and_reopen_releases_cache() {
    let mut host = KernelHost::default();
    open_fixture(&mut host, "a", 7.);
    open_fixture(&mut host, "b", 19.);
    for _ in 0..2 {
        assert_eq!(host.dispatch("analytics.execute", query("a", 1)).unwrap()["rows"][0]["values"], json!([7.]));
        assert_eq!(host.dispatch("analytics.execute", query("b", 1)).unwrap()["rows"][0]["values"], json!([19.]));
    }
    assert_eq!(host.analytics.len(), 2);
    let manifest = host.workbooks["a"].manifest();
    host.dispatch("open", json!({"manifest":manifest})).unwrap();
    assert!(!host.analytics.contains_key("a"));
    assert!(host.analytics.contains_key("b"));
    assert_eq!(host.dispatch("analytics.execute", query("a", 1)).unwrap()["rows"][0]["values"], json!([7.]));
    host.dispatch("close", json!({"unitId":"a"})).unwrap();
    assert!(!host.analytics.contains_key("a"));
    assert_eq!(host.dispatch("analytics.execute", query("a", 1)).unwrap_err().code, "WORKBOOK_NOT_OPEN");
}

#[test]
fn analytics_rejects_stale_and_cancelled_requests_then_recovers() {
    let mut host = KernelHost::default();
    open_fixture(&mut host, "a", 7.);
    let expected = host.dispatch("analytics.execute", query("a", 1)).unwrap();
    assert_eq!(host.dispatch("analytics.execute", query("a", 0)).unwrap_err().code, "STALE_REVISION");
    let bytes = serde_json::to_vec(&json!({"protocolVersion":1,"requestId":"cancel-test",
        "operation":"analytics.execute","params":query("a", 1)})).unwrap();
    let cancelled: Value = serde_json::from_slice(&host.invoke_cancellable(&bytes, Arc::new(AtomicBool::new(true)))).unwrap();
    assert_eq!(cancelled["error"]["code"], "ANALYTICS_TASK_CANCELLED");
    assert!(host.analytics.contains_key("a"));
    assert_eq!(host.dispatch("analytics.execute", query("a", 1)).unwrap(), expected);
}

#[test]
fn workbook_context_capacity_is_bounded_and_close_reclaims_it() {
    let mut host = KernelHost::default();
    for index in 0..MAX_CONTEXTS {
        host.dispatch(
            "create",
            json!({"unitId":format!("unit-{index}"),"name":format!("Unit {index}")}),
        )
        .unwrap();
    }
    let invalid = host
        .dispatch(
            "create",
            json!({"unitId":"invalid","name":"Invalid","sheets":[]}),
        )
        .unwrap_err();
    assert_eq!(invalid.code, "SHEET_INVALID");
    assert!(!host.workbooks.contains_key("invalid"));

    let rejected = host
        .dispatch("create", json!({"unitId":"overflow","name":"Overflow"}))
        .unwrap_err();
    assert_eq!(rejected.code, "KERNEL_CONTEXT_LIMIT");
    assert!(!host.workbooks.contains_key("overflow"));

    host.dispatch("close", json!({"unitId":"unit-0"})).unwrap();
    host.dispatch("create", json!({"unitId":"overflow","name":"Overflow"}))
        .unwrap();
    assert_eq!(host.workbooks.len(), MAX_CONTEXTS);
}
