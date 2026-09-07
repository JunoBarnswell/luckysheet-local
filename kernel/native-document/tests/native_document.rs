use kernel_core::{Cell, CellAddress, CellReader, CellWrite, Scalar, SheetManifest, WorkbookPages};
use kernel_native_document::{Artifact, CodecRevision, NativeDocument, ResourceLimits};
use serde_json::json;
use std::{
    collections::BTreeMap,
    fs::{self, File},
    io::{Read, Write},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};
struct TestDirectory(PathBuf);
impl TestDirectory {
    fn new() -> Self {
        static NEXT: AtomicU64 = AtomicU64::new(0);
        let path = std::env::temp_dir().join(format!(
            "native-document-test-{}-{}",
            std::process::id(),
            NEXT.fetch_add(1, Ordering::Relaxed)
        ));
        fs::create_dir(&path).unwrap();
        Self(path)
    }
}
impl Drop for TestDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}
fn fixture(path: &Path, root_relationship: bool) {
    let mut zip = ZipWriter::new(File::create(path).unwrap());
    let options = SimpleFileOptions::default();
    let mut parts = vec![
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/payload/main.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>"#,
        ),
        (
            "payload/main.xml",
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr date1904="1"/><sheets><sheet name="Data" sheetId="1" r:id="grid"/></sheets></workbook>"#,
        ),
        (
            "payload/_rels/main.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="grid" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="../grids/data.xml"/><Relationship Id="strings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="../dict/words.xml"/><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="../design/formats.xml"/></Relationships>"#,
        ),
        (
            "grids/data.xml",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fut="urn:future"><sheetData><row r="1"><c r="A1" t="s" s="1"><v>0</v><fut:tag keep="yes"/></c><c r="B1"><v>3</v></c></row></sheetData><extLst><fut:payload value="untouched"/></extLst></worksheet>"#,
        ),
        (
            "dict/words.xml",
            r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>A &amp; B</t></si></sst>"#,
        ),
        (
            "design/formats.xml",
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/><sz val="12"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/><xf fontId="1" fillId="0" borderId="0" numFmtId="0"/></cellXfs></styleSheet>"#,
        ),
        ("opaque/custom.bin", "arbitrary unknown bytes"),
        ("code/vbaProject.bin", "macro content never executed"),
    ];
    if root_relationship {
        parts.push(("_rels/.rels",r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="payload/main.xml"/></Relationships>"#));
    }
    for (name, data) in parts {
        zip.start_file(name, options).unwrap();
        zip.write_all(data.as_bytes()).unwrap();
    }
    zip.finish().unwrap();
}
fn part(path: &Path, name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(File::open(path).unwrap()).unwrap();
    let mut bytes = Vec::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_end(&mut bytes)
        .unwrap();
    bytes
}
#[test]
fn source_identity_is_file_bound() {
    let dir = TestDirectory::new();
    let path = dir.0.join("source.xlsx");
    fs::write(&path, b"source").unwrap();
    let artifact = Artifact::from_path(&path, 7, "xlsx".into(), CodecRevision(1)).unwrap();
    artifact.verify().unwrap();
    fs::write(path, b"changed").unwrap();
    assert_eq!(
        artifact.verify().unwrap_err().code,
        "ARTIFACT_CHECKSUM_MISMATCH"
    );
}
#[test]
fn missing_root_relationship_is_rejected() {
    let dir = TestDirectory::new();
    let source = dir.0.join("source.xlsm");
    fixture(&source, false);
    assert!(NativeDocument::open(source, 0, ResourceLimits::default()).is_err());
}
#[test]
fn nonstandard_opc_paths_stream_owned_edits_and_preserve_unknown() {
    let dir = TestDirectory::new();
    let source = dir.0.join("misleading.xlsx");
    fixture(&source, true);
    let doc = NativeDocument::open(&source, 0, ResourceLimits::default()).unwrap();
    assert_eq!(doc.format.as_str(), "xlsm");
    assert_eq!(
        serde_json::to_value(doc.metadata.date_system).unwrap(),
        json!("excel1904")
    );
    let sheets = doc
        .sheets
        .iter()
        .map(|s| SheetManifest {
            sheet_id: s.id.clone(),
            name: s.name.clone(),
            row_count: 100,
            column_count: 20,
            metadata: s.metadata.clone(),
        })
        .collect();
    let mut pages = WorkbookPages::create("unit", "Native", sheets).unwrap();
    let mut writes = Vec::new();
    doc.visit_cells(&dir.0, &mut |r| {
        writes.push(CellWrite {
            address: r.address,
            cell: Some(r.cell),
        });
        Ok(())
    })
    .unwrap();
    assert_eq!(
        writes[0].cell.as_ref().unwrap().value,
        Scalar::Text("A & B".into())
    );
    assert_eq!(
        writes[0].cell.as_ref().unwrap().metadata["style"]["fontSizePx"],
        json!(16.0)
    );
    pages.apply_writes("import", 0, writes).unwrap();
    let address = CellAddress {
        sheet_id: "1".into(),
        row: 0,
        column: 1,
    };
    let mut changed = pages.read_cell(&address).unwrap().unwrap();
    changed.value = Scalar::Text("new & text".into());
    changed.metadata.insert(
        "style".into(),
        json!({"bold":true,"background":"#00FF00","numberFormat":"0.00"}),
    );
    pages
        .apply_writes(
            "edit",
            1,
            vec![
                CellWrite {
                    address,
                    cell: Some(changed),
                },
                CellWrite {
                    address: CellAddress {
                        sheet_id: "1".into(),
                        row: 4,
                        column: 2,
                    },
                    cell: Some(Cell {
                        value: Scalar::Number(42.0),
                        formula: None,
                        metadata: BTreeMap::new(),
                    }),
                },
            ],
        )
        .unwrap();
    let output = dir.0.join("saved.xlsm");
    doc.export_to_path(
        &output,
        0,
        &doc.artifact.identity.checksum,
        &pages.manifest(),
        &pages,
        &dir.0,
    )
    .unwrap();
    for name in ["opaque/custom.bin", "code/vbaProject.bin"] {
        assert_eq!(part(&source, name), part(&output, name));
    }
    let worksheet = String::from_utf8(part(&output, "grids/data.xml")).unwrap();
    assert!(worksheet.contains("<fut:tag keep=\"yes\"/>"));
    assert!(worksheet.contains("<fut:payload value=\"untouched\"/>"));
    let reopened = NativeDocument::open(&output, 2, ResourceLimits::default()).unwrap();
    let mut cells = Vec::new();
    reopened
        .visit_cells(&dir.0, &mut |cell| {
            cells.push(cell);
            Ok(())
        })
        .unwrap();
    assert!(cells.iter().any(|c| c.address.row == 4
        && c.address.column == 2
        && c.cell.value == Scalar::Number(42.0)));
    assert!(cells.iter().any(|c| c.address.column == 1
        && c.cell.value == Scalar::Text("new & text".into())
        && c.cell.metadata["style"]["background"] == json!("#00FF00")));
}
#[test]
fn new_workbook_exports_real_ooxml_and_rejects_stale_binding() {
    let dir = TestDirectory::new();
    let mut pages = WorkbookPages::create(
        "new",
        "Workbook",
        vec![SheetManifest {
            sheet_id: "stable-sheet-id".into(),
            name: "Budget".into(),
            row_count: 100,
            column_count: 20,
            metadata: BTreeMap::new(),
        }],
    )
    .unwrap();
    pages
        .apply_writes(
            "write",
            0,
            vec![CellWrite {
                address: CellAddress {
                    sheet_id: "stable-sheet-id".into(),
                    row: 2,
                    column: 3,
                },
                cell: Some(Cell {
                    value: Scalar::Number(8.0),
                    formula: None,
                    metadata: BTreeMap::new(),
                }),
            }],
        )
        .unwrap();
    let source = dir.0.join("empty.xlsx");
    let doc = NativeDocument::create_source(&source, &pages.manifest(), ResourceLimits::default())
        .unwrap();
    let output = dir.0.join("saved.xlsx");
    assert_eq!(
        doc.export_to_path(&output, 1, "wrong", &pages.manifest(), &pages, &dir.0)
            .unwrap_err()
            .code,
        "ARTIFACT_REVISION_CONFLICT"
    );
    assert!(!output.exists());
    doc.export_to_path(
        &output,
        1,
        &doc.artifact.identity.checksum,
        &pages.manifest(),
        &pages,
        &dir.0,
    )
    .unwrap();
    let imported = NativeDocument::open(&output, 1, ResourceLimits::default()).unwrap();
    let mut count = 0;
    imported
        .visit_cells(&dir.0, &mut |cell| {
            assert_eq!(cell.cell.value, Scalar::Number(8.0));
            count += 1;
            Ok(())
        })
        .unwrap();
    assert_eq!(count, 1);
}
#[test]
fn input_budget_rejects_before_projection() {
    let dir = TestDirectory::new();
    let source = dir.0.join("source.xlsm");
    fixture(&source, true);
    let mut limits = ResourceLimits::default();
    limits.max_archive_bytes = 1;
    assert_eq!(
        NativeDocument::open(source, 0, limits).unwrap_err().code,
        "ZIP_BYTES"
    );
}
