use super::*;
use serde_json::json;

#[test]
fn data_validation_is_rechecked_by_rust() {
    let mut metadata = BTreeMap::new();
    metadata.insert("dataValidations".into(),json!([{"id":"whole","type":"whole","operator":"between","formula1":1,"formula2":10,"allowBlank":false,"alertStyle":"stop","ranges":[{"sheetId":"s","startRow":0,"endRow":9,"startColumn":0,"endColumn":0}]}]));
    let mut pages = WorkbookPages::create(
        "w",
        "Workbook",
        vec![SheetManifest {
            sheet_id: "s".into(),
            name: "Sheet1".into(),
            row_count: 10,
            column_count: 4,
            metadata,
        }],
    )
    .unwrap();
    execute(&mut pages, request(0, "valid", json!([set(0, 5)]))).unwrap();
    let error = execute(&mut pages, request(1, "invalid", json!([set(0, 99)]))).unwrap_err();
    assert_eq!(error.code, "CELL_ENTRY_VALIDATION_FAILED");
    assert_eq!(pages.revision(), 1);
    assert_eq!(
        pages
            .read_cell(&CellAddress {
                sheet_id: "s".into(),
                row: 0,
                column: 0
            })
            .unwrap()
            .unwrap()
            .value,
        Scalar::Number(5.0)
    );
}

#[test]
fn preparation_selects_target_page_without_loading_it() {
    let mut pages = WorkbookPages::create(
        "w",
        "Workbook",
        vec![SheetManifest {
            sheet_id: "s".into(),
            name: "Sheet1".into(),
            row_count: 3000,
            column_count: 4,
            metadata: BTreeMap::new(),
        }],
    )
    .unwrap();
    execute(
        &mut pages,
        request(0, "seed", json!([set(0, 1), set(1500, 2)])),
    )
    .unwrap();
    let unloaded = WorkbookPages::open(pages.manifest()).unwrap();
    let keys = required_pages(&unloaded, &request(1, "edit", json!([set(0, 3)]))).unwrap();
    assert_eq!(
        keys,
        vec![PageKey {
            sheet_id: "s".into(),
            page_row: 0,
            page_column: 0
        }]
    );
    assert_eq!(
        required_pages(&unloaded, &request(0, "stale", json!([set(0, 3)])))
            .unwrap_err()
            .code,
        "STALE_REVISION"
    );
}

fn book() -> WorkbookPages {
    WorkbookPages::create(
        "w",
        "Workbook",
        vec![SheetManifest {
            sheet_id: "s".into(),
            name: "Sheet1".into(),
            row_count: 10,
            column_count: 4,
            metadata: BTreeMap::new(),
        }],
    )
    .unwrap()
}
fn request(revision: u64, id: &str, mutations: Value) -> CommandRequest {
    CommandRequest {
        unit_id: "w".into(),
        base_revision: revision,
        operation_id: id.into(),
        command_id: "operation.apply".into(),
        params: json!({"mutations":mutations}),
    }
}
fn set(row: u32, value: i32) -> Value {
    json!({"id":"cell.set","sheetId":"s","params":{"sheetId":"s","row":row,"column":0,"value":{"value":value}}})
}
fn undo(revision: u64, id: &str, history: &HistoryRecord) -> CommandRequest {
    CommandRequest {
        unit_id: "w".into(), base_revision: revision, operation_id: id.into(),
        command_id: "history.undo".into(), params: json!({"history":history}),
    }
}

#[test]
fn undo_reverts_owned_pages_in_a_new_revision_and_requires_editor() {
    let mut pages = book();
    let committed = execute(&mut pages, request(0, "target", json!([set(0, 42)]))).unwrap();
    assert!(required_pages(&pages, &undo(1, "undo", &committed.history)).unwrap().is_empty());
    assert_eq!(execute_authorized(&mut pages, undo(1, "denied", &committed.history), AccessRole::Commenter).unwrap_err().code, "FORBIDDEN");
    let reverted = execute(&mut pages, undo(1, "undo", &committed.history)).unwrap();
    assert_eq!(reverted.revision, 2);
    assert_eq!(reverted.history.page_deltas[0].before, committed.history.page_deltas[0].after);
    assert_eq!(reverted.history.page_deltas[0].after, committed.history.page_deltas[0].before);
    assert_eq!(pages.read_cell(&CellAddress { sheet_id: "s".into(), row: 0, column: 0 }).unwrap(), None);
}

#[test]
fn undo_preserves_nonoverlapping_pages_and_rejects_overlap_atomically() {
    let mut pages = WorkbookPages::create("w", "Workbook", vec![SheetManifest {
        sheet_id: "s".into(), name: "Sheet1".into(), row_count: 3000, column_count: 4, metadata: BTreeMap::new(),
    }]).unwrap();
    let target = execute(&mut pages, request(0, "target", json!([set(0, 42)]))).unwrap();
    execute(&mut pages, request(1, "other-page", json!([set(1500, 9)]))).unwrap();
    execute(&mut pages, undo(2, "undo", &target.history)).unwrap();
    assert_eq!(pages.read_cell(&CellAddress { sheet_id: "s".into(), row: 1500, column: 0 }).unwrap().unwrap().value, Scalar::Number(9.));

    let mut conflicting = book();
    let target = execute(&mut conflicting, request(0, "target", json!([set(0, 42)]))).unwrap();
    execute(&mut conflicting, request(1, "overlap", json!([set(0, 99)]))).unwrap();
    assert_eq!(execute(&mut conflicting, undo(2, "undo", &target.history)).unwrap_err().code, "UNDO_CONFLICT");
    assert_eq!(conflicting.revision(), 2);
    assert_eq!(conflicting.read_cell(&CellAddress { sheet_id: "s".into(), row: 0, column: 0 }).unwrap().unwrap().value, Scalar::Number(99.));
}

#[test]
fn undo_metadata_requires_the_recorded_after_state() {
    let rename = |name: &str| json!([{"id":"sheet.rename","sheetId":"s","params":{"sheetId":"s","name":name}}]);
    let mut pages = book();
    let target = execute(&mut pages, request(0, "rename", rename("Renamed"))).unwrap();
    assert_eq!(target.history.metadata_after.as_ref().unwrap().sheets[0].name, "Renamed");
    execute(&mut pages, undo(1, "undo-rename", &target.history)).unwrap();
    assert_eq!(pages.manifest().sheets[0].name, "Sheet1");

    let mut conflicting = book();
    let target = execute(&mut conflicting, request(0, "rename", rename("Renamed"))).unwrap();
    execute(&mut conflicting, request(1, "rename-again", rename("Later"))).unwrap();
    assert_eq!(execute(&mut conflicting, undo(2, "undo-rename", &target.history)).unwrap_err().code, "UNDO_CONFLICT");
    assert_eq!(conflicting.manifest().sheets[0].name, "Later");
}

#[test]
fn batch_has_one_revision_and_roundtrips_pages() {
    let mut pages = book();
    let result = execute(
        &mut pages,
        request(0, "op1", json!([set(0, 12), set(1, 30)])),
    )
    .unwrap();
    assert_eq!(result.revision, 1);
    assert_eq!(result.pages.len(), 1);
    assert_eq!(
        pages
            .read_cell(&CellAddress {
                sheet_id: "s".into(),
                row: 1,
                column: 0
            })
            .unwrap()
            .unwrap()
            .value,
        Scalar::Number(30.0)
    );
    let mut reopened = WorkbookPages::open(result.manifest.clone()).unwrap();
    for payload in result.pages {
        let bytes = pages.page_bytes(&payload.descriptor.key()).unwrap();
        reopened.load_page(&payload.descriptor, &bytes).unwrap();
    }
    assert_eq!(
        reopened
            .read_cell(&CellAddress {
                sheet_id: "s".into(),
                row: 0,
                column: 0
            })
            .unwrap()
            .unwrap()
            .value,
        Scalar::Number(12.0)
    );
}
#[test]
fn invalid_tail_rejects_without_partial_cell_or_revision() {
    let mut pages = book();
    let error = execute(
        &mut pages,
        request(
            0,
            "bad",
            json!([set(0,12),{"id":"not.a.command","sheetId":"s","params":{}}]),
        ),
    )
    .unwrap_err();
    assert_eq!(error.code, "COMMAND_UNKNOWN");
    assert_eq!(pages.revision(), 0);
    assert_eq!(
        pages
            .read_cell(&CellAddress {
                sheet_id: "s".into(),
                row: 0,
                column: 0
            })
            .unwrap(),
        None
    );
}
#[test]
fn stale_revision_and_role_are_rejected() {
    let mut pages = book();
    let error = execute(&mut pages, request(3, "stale", json!([set(0, 12)]))).unwrap_err();
    assert_eq!(error.code, "STALE_REVISION");
    let error = execute_authorized(
        &mut pages,
        request(0, "forbidden", json!([set(0, 12)])),
        AccessRole::Commenter,
    )
    .unwrap_err();
    assert_eq!(error.code, "FORBIDDEN");
    assert_eq!(pages.revision(), 0);
}
#[test]
fn protected_cells_reject_editor_but_owner_can_commit() {
    let mut pages = book();
    let protection = json!({"id":"sheet.protect.set","sheetId":"s","params":{"rule":{"id":"locked","scope":"sheet","sheetId":"s","locked":true,"allow":{}}}});
    execute_authorized(
        &mut pages,
        request(0, "protect", json!([protection])),
        AccessRole::Owner,
    )
    .unwrap();
    let error = execute(&mut pages, request(1, "blocked", json!([set(0, 12)]))).unwrap_err();
    assert_eq!(error.code, "FORBIDDEN");
    assert_eq!(pages.revision(), 1);
    execute_authorized(
        &mut pages,
        request(1, "owner", json!([set(0, 12)])),
        AccessRole::Owner,
    )
    .unwrap();
    assert_eq!(pages.revision(), 2);
}
#[test]
fn grow_then_write_is_one_transaction() {
    let mut pages = book();
    let grow = json!({"id":"sheet.extent.grow","sheetId":"s","params":{"sheetId":"s","rowCount":20,"columnCount":4}});
    let result = execute(&mut pages, request(0, "grow", json!([grow, set(19, 99)]))).unwrap();
    assert_eq!(result.revision, 1);
    assert_eq!(result.manifest.sheets[0].row_count, 20);
}
