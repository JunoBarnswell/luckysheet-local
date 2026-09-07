use crate::{
    archive::Package,
    artifact::Artifact,
    shared_strings::SharedStrings,
    styles::Styles,
    worksheet, xml,
    xmlnode::{self, error, Node},
};
use kernel_core::{
    Cell, CellAddress, CellReader, KernelError, KernelResult, SheetManifest, WorkbookManifest,
    MAX_COLUMNS, MAX_ROWS,
};
use quick_xml::{events::Event, Reader};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::{
    collections::{BTreeMap, BTreeSet},
    fs::{File, OpenOptions},
    io::{BufWriter, Read, Write},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum DateSystem {
    Excel1900,
    Excel1904,
}
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum DocumentFormat {
    Xlsx,
    Xlsm,
    Xltx,
    Xltm,
    Xlam,
}
impl DocumentFormat {
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::Xlsx => "xlsx",
            Self::Xlsm => "xlsm",
            Self::Xltx => "xltx",
            Self::Xltm => "xltm",
            Self::Xlam => "xlam",
        }
    }
}
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum FormatSupport {
    Edit,
    Preserve,
    ExportOnly,
    Unsupported,
}
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct FeatureCapability {
    pub feature: String,
    pub support: FormatSupport,
    pub reason: String,
}
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields, default)]
pub struct ResourceLimits {
    pub max_archive_bytes: u64,
    pub max_entries: u64,
    pub max_entry_bytes: u64,
    pub max_uncompressed_bytes: u64,
    pub max_temporary_bytes: u64,
    pub max_cells: u64,
    pub max_metadata_bytes: u64,
    pub max_cell_bytes: u64,
    pub max_xml_depth: u32,
    pub max_working_bytes: u64,
}
impl Default for ResourceLimits {
    fn default() -> Self {
        Self {
            max_archive_bytes: 1024 * 1024 * 1024,
            max_entries: 20_000,
            max_entry_bytes: 4 * 1024 * 1024 * 1024,
            max_uncompressed_bytes: 8 * 1024 * 1024 * 1024,
            max_temporary_bytes: 16 * 1024 * 1024 * 1024,
            max_cells: 20_000_000,
            max_metadata_bytes: 16 * 1024 * 1024,
            max_cell_bytes: 1024 * 1024,
            max_xml_depth: 128,
            max_working_bytes: 256 * 1024 * 1024,
        }
    }
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct CellRecord {
    pub address: CellAddress,
    pub cell: Cell,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SheetMetadata {
    pub id: String,
    pub name: String,
    pub part: String,
    pub hidden: bool,
    pub metadata: BTreeMap<String, Value>,
}
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Metadata {
    pub date_system: DateSystem,
    pub features: Vec<FeatureCapability>,
    pub workbook_metadata: BTreeMap<String, Value>,
}
#[derive(Debug, Clone)]
struct Relationship {
    id: String,
    kind: String,
    target: String,
    external: bool,
}
#[derive(Debug, Clone)]
pub struct NativeDocument {
    pub format: DocumentFormat,
    pub metadata: Metadata,
    pub sheets: Vec<SheetMetadata>,
    pub artifact: Artifact,
    package: Package,
    workbook_part: String,
    workbook: Vec<u8>,
    styles_part: Option<String>,
    styles: Styles,
    shared_strings_part: Option<String>,
    limits: ResourceLimits,
}

impl NativeDocument {
    /// Creates the real OPC skeleton for a workbook which has no imported source.
    /// Canonical cells are then streamed through the same owned worksheet writer.
    pub fn create_source(
        path: impl AsRef<Path>,
        manifest: &WorkbookManifest,
        limits: ResourceLimits,
    ) -> KernelResult<Self> {
        use zip::{write::SimpleFileOptions, ZipWriter};
        kernel_core::WorkbookPages::open(manifest.clone())?;
        let path = path.as_ref();
        let file = OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(path)
            .map_err(io_error)?;
        let result = (|| {
            let mut zip = ZipWriter::new(file);
            let options =
                SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
            let mut types=String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            let mut workbook=String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr date1904=\"0\"/><sheets>");
            let mut rels=String::from("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"styles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            for (index, sheet) in manifest.sheets.iter().enumerate() {
                let id = index + 1;
                let part = format!("xl/worksheets/sheet{id}.xml");
                types.push_str(&format!("<Override PartName=\"/{part}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"));
                workbook.push_str(&format!(
                    "<sheet sheetId=\"{id}\" name=\"{}\" r:id=\"sheet{id}\"/>",
                    xmlnode::escape(&sheet.name)
                ));
                rels.push_str(&format!("<Relationship Id=\"sheet{id}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{id}.xml\"/>"));
                zip.start_file(part, options).map_err(io_error)?;
                zip.write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>").map_err(io_error)?;
            }
            types.push_str("</Types>");
            workbook.push_str("</sheets><calcPr fullCalcOnLoad=\"1\"/></workbook>");
            rels.push_str("</Relationships>");
            let styles = Styles::default_styles()?.bytes()?;
            for(name,bytes)in [("[Content_Types].xml",types.as_bytes()),("_rels/.rels",b"<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"workbook\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>".as_slice()),("xl/workbook.xml",workbook.as_bytes()),("xl/_rels/workbook.xml.rels",rels.as_bytes()),("xl/styles.xml",styles.as_slice())]{zip.start_file(name,options).map_err(io_error)?;zip.write_all(bytes).map_err(io_error)?;}
            zip.finish()
                .map_err(io_error)?
                .sync_all()
                .map_err(io_error)?;
            Ok(())
        })();
        if let Err(e) = result {
            let _ = std::fs::remove_file(path);
            return Err(e);
        }
        let mut document = Self::open(path, manifest.revision, limits)?;
        for (source, sheet) in document.sheets.iter_mut().zip(&manifest.sheets) {
            source.id = sheet.sheet_id.clone();
        }
        Ok(document)
    }
    pub fn open(
        path: impl AsRef<Path>,
        revision: u64,
        limits: ResourceLimits,
    ) -> KernelResult<Self> {
        let package = Package::open(path.as_ref(), limits.clone())?;
        let parts: BTreeSet<_> = package.part_names()?.into_iter().collect();
        let root_rels = relationships(&package, None, &limits)?;
        let workbook_relation =
            unique_relationship(&root_rels, "officeDocument")?.ok_or_else(|| {
                error(
                    "OOXML_WORKBOOK_MISSING",
                    "OPC root officeDocument relationship is required",
                )
            })?;
        if workbook_relation.external {
            return Err(error(
                "OOXML_WORKBOOK_EXTERNAL",
                "Workbook must be an internal OPC part",
            ));
        }
        let workbook_part = workbook_relation.target.clone();
        let content_types = xmlnode::parse(
            &read_metadata(&package, "[Content_Types].xml", &limits)?,
            limits.max_metadata_bytes,
        )?;
        let content_type = content_types
            .children_named("Override")
            .find(|n| {
                n.attr("PartName").map(|p| p.trim_start_matches('/'))
                    == Some(workbook_part.as_str())
            })
            .and_then(|n| n.attr("ContentType"))
            .ok_or_else(|| error("OOXML_CONTENT_TYPE_MISSING", workbook_part.clone()))?;
        let format=match content_type{"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"=>DocumentFormat::Xlsx,"application/vnd.ms-excel.sheet.macroEnabled.main+xml"=>DocumentFormat::Xlsm,"application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"=>DocumentFormat::Xltx,"application/vnd.ms-excel.template.macroEnabled.main+xml"=>DocumentFormat::Xltm,"application/vnd.ms-excel.addin.macroEnabled.main+xml"=>DocumentFormat::Xlam,_=>return Err(error("UNSUPPORTED_FORMAT",format!("Workbook content type {content_type} is not an editable OOXML worksheet document")))};
        let workbook = read_metadata(&package, &workbook_part, &limits)?;
        let root = xmlnode::parse(&workbook, limits.max_metadata_bytes)?;
        if root.local() != "workbook" {
            return Err(error("OOXML_WORKBOOK_INVALID", workbook_part.clone()));
        }
        let rels = relationships(&package, Some(&workbook_part), &limits)?;
        let mut sheets = Vec::new();
        let mut ids = BTreeSet::new();
        let mut names = BTreeSet::new();
        let parent = root
            .child("sheets")
            .ok_or_else(|| error("OOXML_SHEETS_MISSING", "Workbook has no sheets collection"))?;
        for sheet in parent.children_named("sheet") {
            let id = required(sheet, "sheetId")?.to_owned();
            let name = required(sheet, "name")?.to_owned();
            if !ids.insert(id.clone()) || !names.insert(name.to_lowercase()) {
                return Err(error("OOXML_SHEET_DUPLICATE", name));
            }
            let rid = sheet
                .attributes
                .iter()
                .find(|(k, _)| k.rsplit(':').next() == Some("id"))
                .map(|(_, v)| v.as_str())
                .ok_or_else(|| error("OOXML_RELATIONSHIP_MISSING", name.clone()))?;
            let rel = rels
                .iter()
                .find(|r| r.id == rid)
                .ok_or_else(|| error("OOXML_RELATIONSHIP_MISSING", rid))?;
            if rel.external || !relationship_kind(&rel.kind, "worksheet") {
                return Err(error(
                    "UNSUPPORTED_FEATURE",
                    format!("Sheet {name} has unsupported relationship {}", rel.kind),
                ));
            }
            if !parts.contains(&rel.target) {
                return Err(error("OOXML_SHEET_MISSING", rel.target.clone()));
            }
            let state = sheet.attr("state").unwrap_or("visible");
            if !["visible", "hidden", "veryHidden"].contains(&state) {
                return Err(error("OOXML_SHEET_STATE_INVALID", state));
            }
            let mut metadata = sheet_metadata(&package, &rel.target, &id, &limits)?;
            metadata.insert("hidden".into(), json!(state != "visible"));
            if state == "veryHidden" {
                metadata.insert("nativeVeryHidden".into(), json!(true));
            }
            sheets.push(SheetMetadata {
                id,
                name,
                part: rel.target.clone(),
                hidden: state != "visible",
                metadata,
            });
        }
        if sheets.is_empty() {
            return Err(error("OOXML_SHEETS_MISSING", "Workbook has no worksheets"));
        }
        let date_system = match root.child("workbookPr").and_then(|n| n.attr("date1904")) {
            None | Some("0" | "false") => DateSystem::Excel1900,
            Some("1" | "true") => DateSystem::Excel1904,
            Some(v) => return Err(error("OOXML_DATE_SYSTEM_INVALID", v)),
        };
        let styles_part = unique_relationship(&rels, "styles")?.map(|r| r.target.clone());
        let theme = unique_relationship(&rels, "theme")?
            .map(|r| read_metadata(&package, &r.target, &limits))
            .transpose()?;
        let styles = match &styles_part {
            Some(part) => Styles::parse(
                read_metadata(&package, part, &limits)?,
                theme.as_deref(),
                limits.max_metadata_bytes,
            )?,
            None => Styles::default_styles()?,
        };
        let shared_strings_part =
            unique_relationship(&rels, "sharedStrings")?.map(|r| r.target.clone());
        let mut workbook_metadata = BTreeMap::new();
        workbook_metadata.insert(
            "dateSystem".into(),
            json!(match date_system {
                DateSystem::Excel1900 => "1900",
                DateSystem::Excel1904 => "1904",
            }),
        );
        if let Some(names) = root.child("definedNames") {
            let mut definitions = Vec::new();
            for n in names.children_named("definedName") {
                let name = required(n, "name")?;
                let scope = n
                    .attr("localSheetId")
                    .map(|v| {
                        v.parse::<usize>()
                            .map_err(|_| error("OOXML_NAME_SCOPE_INVALID", v))
                    })
                    .transpose()?;
                let sheet_id = scope
                    .map(|i| {
                        sheets
                            .get(i)
                            .map(|s| s.id.clone())
                            .ok_or_else(|| error("OOXML_NAME_SCOPE_INVALID", i.to_string()))
                    })
                    .transpose()?;
                let mut definition = json!({"name":name,"formula":format!("={}",n.text),"scope":if sheet_id.is_some(){"sheet"}else{"workbook"}});
                if let Some(sheet_id) = sheet_id {
                    definition["sheetId"] = json!(sheet_id);
                }
                if let Some(hidden) = n.attr("hidden") {
                    definition["hidden"] = json!(hidden == "1" || hidden == "true");
                }
                if let Some(comment) = n.attr("comment") {
                    definition["comment"] = json!(comment);
                }
                definitions.push(definition);
            }
            workbook_metadata.insert("definedNameModels".into(), json!(definitions));
        }
        let features = features(&parts, &rels);
        workbook_metadata.insert("nativeCapabilities".into(), json!(features));
        let artifact = Artifact::from_path(
            path,
            revision,
            format.as_str().into(),
            crate::NATIVE_DOCUMENT_CODEC_REVISION,
        )?;
        Ok(Self {
            format,
            metadata: Metadata {
                date_system,
                features,
                workbook_metadata,
            },
            sheets,
            artifact,
            package,
            workbook_part,
            workbook,
            styles_part,
            styles,
            shared_strings_part,
            limits,
        })
    }
    /// The visitor receives row-ordered cells and may commit each completed page
    /// stripe directly to a task file. No document-sized cell collection exists.
    pub fn visit_cells(
        &self,
        temporary_directory: impl AsRef<Path>,
        visitor: &mut dyn FnMut(CellRecord) -> KernelResult<()>,
    ) -> KernelResult<u64> {
        self.artifact.verify()?;
        let mut shared = SharedStrings::build(
            &self.package,
            self.shared_strings_part.as_deref(),
            temporary_directory.as_ref(),
            &self.limits,
        )?;
        let mut count = 0u64;
        for sheet in &self.sheets {
            count += self.package.with_part_reader(&sheet.part, |reader| {
                worksheet::parse_sheet(
                    &sheet.id,
                    reader,
                    &mut |i| shared.get(i),
                    &self.styles.values,
                    &self.limits,
                    visitor,
                )
            })?;
            if count > self.limits.max_cells {
                return Err(error(
                    "CELL_LIMIT",
                    "Document cell count exceeds the task budget",
                ));
            }
        }
        Ok(count)
    }
    pub fn export_to_path(
        &self,
        output: impl AsRef<Path>,
        source_revision: u64,
        source_checksum: &str,
        manifest: &WorkbookManifest,
        cells: &dyn CellReader,
        temporary_directory: impl AsRef<Path>,
    ) -> KernelResult<Artifact> {
        self.artifact.verify()?;
        if source_revision != self.artifact.identity.revision
            || source_checksum != self.artifact.identity.checksum
        {
            return Err(error(
                "ARTIFACT_REVISION_CONFLICT",
                "Source revision/checksum does not match the bound native artifact",
            ));
        }
        if manifest.revision != cells.revision() || manifest.revision < source_revision {
            return Err(error(
                "REVISION_CONFLICT",
                "Export cells and manifest must belong to the target revision",
            ));
        }
        for key in ["dateSystem", "definedNameModels", "printDocuments"] {
            if manifest.metadata.contains_key(key)
                && !metadata_equal(
                    manifest.metadata.get(key),
                    self.metadata.workbook_metadata.get(key),
                )
            {
                return Err(error(
                    "UNSUPPORTED_FEATURE",
                    format!("Workbook metadata {key} changed without an owned native writer"),
                ));
            }
        }
        if manifest.sheets.len() != self.sheets.len()
            || manifest
                .sheets
                .iter()
                .zip(&self.sheets)
                .any(|(a, b)| a.sheet_id != b.id)
        {
            return Err(error(
                "UNSUPPORTED_FEATURE",
                "Worksheet creation/deletion/reordering requires an owned OPC graph transaction",
            ));
        }
        let parts = self.package.part_names()?;
        if parts.iter().any(|p| p.starts_with("_xmlsignatures/")) {
            return Err(error(
                "UNSUPPORTED_FEATURE",
                "Editing a signed package requires explicit signature removal ownership",
            ));
        }
        for sheet in &self.sheets {
            let current = manifest
                .sheets
                .iter()
                .find(|s| s.sheet_id == sheet.id)
                .unwrap();
            for (key, value) in &sheet.metadata {
                if !["hidden"].contains(&key.as_str())
                    && !metadata_equal(current.metadata.get(key), Some(value))
                {
                    return Err(error(
                        "UNSUPPORTED_FEATURE",
                        format!(
                            "Worksheet {} metadata {key} has no owned OOXML writer",
                            sheet.id
                        ),
                    ));
                }
            }
        }
        for current in &manifest.sheets {
            let original = self
                .sheets
                .iter()
                .find(|s| s.id == current.sheet_id)
                .unwrap();
            for key in [
                "merges",
                "pane",
                "rowHeightsPx",
                "columnWidthsPx",
                "hiddenRows",
                "hiddenColumns",
                "sheetTables",
                "autoFilter",
                "pivots",
                "drawings",
                "drawingPayloads",
                "review",
                "conditionalFormats",
                "dataValidations",
                "sparklines",
                "sparklineGroups",
                "protectionRules",
                "outline",
            ] {
                if !metadata_equal(current.metadata.get(key), original.metadata.get(key)) {
                    return Err(error("UNSUPPORTED_FEATURE",format!("Worksheet {} changed {key}; this native writer does not own that structural metadata",current.sheet_id)));
                }
            }
        }
        let directory = temporary_directory.as_ref();
        let mut temporary = TemporaryParts::new();
        let mut replacements = BTreeMap::new();
        let mut shared = SharedStrings::build(
            &self.package,
            self.shared_strings_part.as_deref(),
            directory,
            &self.limits,
        )?;
        let mut styles = self.styles.clone();
        for sheet in &self.sheets {
            let path = temporary.create(directory, "sheet")?;
            let output = File::create(&path).map_err(io_error)?;
            let row_count = manifest
                .pages
                .iter()
                .filter(|p| p.sheet_id == sheet.id)
                .filter_map(|p| p.occupied_range.as_ref())
                .map(|r| r.end_row + 1)
                .max()
                .unwrap_or(0);
            self.package.with_part_reader(&sheet.part, |reader| {
                worksheet::rewrite_sheet(
                    &sheet.id,
                    reader,
                    BufWriter::new(output),
                    cells,
                    row_count,
                    &mut |i| shared.get(i),
                    &self.styles.values,
                    &mut |cell| styles.index_for(cell),
                )
            })?;
            if std::fs::metadata(&path).map_err(io_error)?.len() > self.limits.max_entry_bytes {
                return Err(error(
                    "ZIP_ENTRY_BYTES",
                    "Rewritten worksheet exceeds its part budget",
                ));
            }
            replacements.insert(sheet.part.clone(), path);
        }
        if styles.changed() {
            let part = self.styles_part.as_ref().ok_or_else(|| {
                error(
                    "UNSUPPORTED_FEATURE",
                    "Creating a styles relationship is not owned by this package transaction",
                )
            })?;
            let path = temporary.create(directory, "styles")?;
            std::fs::write(&path, styles.bytes()?).map_err(io_error)?;
            replacements.insert(part.clone(), path);
        }
        let updated_workbook = rewrite_sheet_directory(
            &self.workbook,
            &manifest.sheets,
            &self.sheets,
            self.limits.max_metadata_bytes,
        )?;
        if updated_workbook != self.workbook {
            let path = temporary.create(directory, "workbook")?;
            std::fs::write(&path, updated_workbook).map_err(io_error)?;
            replacements.insert(self.workbook_part.clone(), path);
        }
        let output = output.as_ref();
        if output == self.artifact.path {
            return Err(error(
                "ARTIFACT_OVERWRITE_FORBIDDEN",
                "Export must create a separate transaction artifact",
            ));
        }
        let result = self.package.write_to_path(output, &replacements);
        if let Err(e) = result {
            let _ = std::fs::remove_file(output);
            return Err(e);
        }
        Artifact::from_path(
            output,
            manifest.revision,
            self.format.as_str().into(),
            crate::NATIVE_DOCUMENT_CODEC_REVISION,
        )
    }
}
fn read_metadata(package: &Package, name: &str, limits: &ResourceLimits) -> KernelResult<Vec<u8>> {
    package.with_part_reader(name, |reader| {
        let mut bytes = Vec::new();
        reader
            .take(limits.max_metadata_bytes + 1)
            .read_to_end(&mut bytes)
            .map_err(io_error)?;
        if bytes.len() as u64 > limits.max_metadata_bytes {
            return Err(error("XML_METADATA_BUDGET", name));
        }
        Ok(bytes)
    })
}
fn required<'a>(node: &'a Node, name: &str) -> KernelResult<&'a str> {
    node.attr(name)
        .filter(|v| !v.is_empty())
        .ok_or_else(|| error("OOXML_ATTRIBUTE_MISSING", format!("{}@{name}", node.name)))
}
fn rels_part(source: Option<&str>) -> String {
    match source {
        None => "_rels/.rels".into(),
        Some(s) => match s.rsplit_once('/') {
            Some((p, n)) => format!("{p}/_rels/{n}.rels"),
            None => format!("_rels/{s}.rels"),
        },
    }
}
fn relationships(
    package: &Package,
    source: Option<&str>,
    limits: &ResourceLimits,
) -> KernelResult<Vec<Relationship>> {
    let part = rels_part(source);
    let root = xmlnode::parse(
        &read_metadata(package, &part, limits)?,
        limits.max_metadata_bytes,
    )?;
    if root.local() != "Relationships" {
        return Err(error("OOXML_RELATIONSHIPS_INVALID", part));
    }
    let mut ids = BTreeSet::new();
    let mut rels = Vec::new();
    for n in root.children_named("Relationship") {
        let id = required(n, "Id")?.to_owned();
        if !ids.insert(id.clone()) {
            return Err(error("OOXML_RELATIONSHIP_DUPLICATE", id));
        }
        let kind = required(n, "Type")?.to_owned();
        let target = required(n, "Target")?;
        let external = match n.attr("TargetMode") {
            None | Some("Internal") => false,
            Some("External") => true,
            _ => return Err(error("OOXML_RELATIONSHIP_MODE_INVALID", id)),
        };
        rels.push(Relationship {
            id,
            kind,
            target: if external {
                target.into()
            } else {
                xml::resolve_relationship_target(target, source)?
            },
            external,
        });
    }
    Ok(rels)
}
fn relationship_kind(kind: &str, suffix: &str) -> bool {
    kind == format!("http://schemas.openxmlformats.org/officeDocument/2006/relationships/{suffix}")
        || kind == format!("http://purl.oclc.org/ooxml/officeDocument/relationships/{suffix}")
}
fn unique_relationship<'a>(
    rels: &'a [Relationship],
    kind: &str,
) -> KernelResult<Option<&'a Relationship>> {
    let mut matches = rels.iter().filter(|r| relationship_kind(&r.kind, kind));
    let result = matches.next();
    if matches.next().is_some() {
        return Err(error("OOXML_RELATIONSHIP_DUPLICATE", kind));
    }
    if result.is_some_and(|r| r.external) {
        return Err(error("OOXML_RELATIONSHIP_EXTERNAL", kind));
    }
    Ok(result)
}
fn features(parts: &BTreeSet<String>, rels: &[Relationship]) -> Vec<FeatureCapability> {
    let mut values = vec![
        FeatureCapability {
            feature: "cells".into(),
            support: FormatSupport::Edit,
            reason: "Native streaming scalar/formula/rich-text cell ownership".into(),
        },
        FeatureCapability {
            feature: "styles".into(),
            support: FormatSupport::Edit,
            reason: "Native font/fill/border/alignment/number format records".into(),
        },
    ];
    for (kind, feature) in [("vbaProject", "macro"), ("pivotCacheDefinition", "pivot")] {
        if rels.iter().any(|r| r.kind.ends_with(&format!("/{kind}"))) {
            values.push(FeatureCapability {
                feature: feature.into(),
                support: FormatSupport::Preserve,
                reason: "Original OPC part retained; native execution/editing is not implemented"
                    .into(),
            });
        }
    }
    for feature in ["tables", "charts", "comments", "drawings"] {
        if parts
            .iter()
            .any(|p| p.to_ascii_lowercase().contains(feature))
        {
            values.push(FeatureCapability {
                feature: feature.into(),
                support: FormatSupport::Preserve,
                reason: "Original parts retained; typed semantic editor is not yet implemented"
                    .into(),
            });
        }
    }
    values
}
fn sheet_metadata(
    package: &Package,
    part: &str,
    sheet_id: &str,
    limits: &ResourceLimits,
) -> KernelResult<BTreeMap<String, Value>> {
    package.with_part_reader(part,|input|{
 let mut reader=Reader::from_reader(input);let mut buffer=Vec::new();let mut map=BTreeMap::new();let mut merges=Vec::new();let mut heights=BTreeMap::new();let mut widths=BTreeMap::new();let mut hidden_rows=Vec::new();let mut hidden_columns=Vec::new();
 loop{match reader.read_event_into(&mut buffer).map_err(|e|error("XML_INVALID",e.to_string()))?{
 Event::Start(e)|Event::Empty(e)=>{let key=xml::local_name(e.name().as_ref()).to_vec();let attr=|k:&[u8]|xml::attr(e.attributes(),k);match key.as_slice(){
 b"mergeCell"=>{let reference=attr(b"ref")?.ok_or_else(||error("OOXML_MERGE_INVALID","Merge ref required"))?;let (a,b)=reference.split_once(':').ok_or_else(||error("OOXML_MERGE_INVALID",reference.clone()))?;let (sr,sc)=worksheet::parse_ref(a)?;let(er,ec)=worksheet::parse_ref(b)?;if sr>er||sc>ec{return Err(error("OOXML_MERGE_INVALID",reference));}merges.push(json!({"range":{"sheetId":sheet_id,"startRow":sr,"endRow":er,"startColumn":sc,"endColumn":ec},"anchor":{"row":sr,"column":sc}}));},
 b"sheetFormatPr"=>{if let Some(v)=attr(b"defaultRowHeight")?{let n=v.parse::<f64>().map_err(|_|error("OOXML_DIMENSION_INVALID",v))?;map.insert("defaultRowHeightPx".into(),json!(n*96.0/72.0));}},
 b"row"=>{let Some(r)=attr(b"r")?else{return Err(error("OOXML_ROW_ADDRESS_MISSING",part))};let row=r.parse::<u32>().map_err(|_|error("OOXML_ROW_INVALID",&r))?.checked_sub(1).ok_or_else(||error("OOXML_ROW_INVALID",&r))?;if row>=MAX_ROWS{return Err(error("OOXML_ROW_INVALID",r));}if let Some(v)=attr(b"ht")?{let n=v.parse::<f64>().map_err(|_|error("OOXML_DIMENSION_INVALID",v))?;heights.insert(row,json!(n*96.0/72.0));}if attr(b"hidden")?.is_some_and(|v|v=="1"||v=="true"){hidden_rows.push(row);}},
 b"col"=>{let min=attr(b"min")?.ok_or_else(||error("OOXML_COLUMN_INVALID","min required"))?.parse::<u32>().map_err(|_|error("OOXML_COLUMN_INVALID","min"))?;let max=attr(b"max")?.ok_or_else(||error("OOXML_COLUMN_INVALID","max required"))?.parse::<u32>().map_err(|_|error("OOXML_COLUMN_INVALID","max"))?;if min==0||min>max||max>MAX_COLUMNS{return Err(error("OOXML_COLUMN_INVALID",part));}let width=attr(b"width")?.map(|v|v.parse::<f64>().map_err(|_|error("OOXML_COLUMN_INVALID",v))).transpose()?;let hidden=attr(b"hidden")?.is_some_and(|v|v=="1"||v=="true");for col in min-1..max{if let Some(w)=width{widths.insert(col,json!(w));}if hidden{hidden_columns.push(col);}}},
 b"sheetView"=>{for (native,canonical) in [(b"showGridLines".as_slice(),"showGridlines"),(b"showRowColHeaders".as_slice(),"showHeaders")]{if let Some(v)=attr(native)?{map.insert(canonical.into(),json!(v!="0"&&v!="false"));}}if let Some(v)=attr(b"zoomScale")?{map.insert("zoom".into(),json!(v.parse::<f64>().map_err(|_|error("OOXML_ZOOM_INVALID",v))?/100.0));}},
 b"pane"=>{let state=attr(b"state")?.unwrap_or_else(||"split".into());if !["frozen","frozenSplit","split"].contains(&state.as_str()){return Err(error("OOXML_PANE_INVALID",state));}let x=attr(b"xSplit")?.map(|v|v.parse::<f64>().map_err(|_|error("OOXML_PANE_INVALID",v))).transpose()?.unwrap_or(0.0);let y=attr(b"ySplit")?.map(|v|v.parse::<f64>().map_err(|_|error("OOXML_PANE_INVALID",v))).transpose()?.unwrap_or(0.0);let (row,column)=attr(b"topLeftCell")?.map(|v|worksheet::parse_ref(&v)).transpose()?.unwrap_or((0,0));if !x.is_finite()||!y.is_finite()||x<0.0||y<0.0{return Err(error("OOXML_PANE_INVALID","Invalid split position"));}let mut pane=json!({"kind":if state=="split"{"split"}else{"frozen"},"state":state,"xSplit":x,"ySplit":y,"startRow":row,"startColumn":column});if let Some(active)=attr(b"activePane")?{if !["topLeft","topRight","bottomLeft","bottomRight"].contains(&active.as_str()){return Err(error("OOXML_PANE_INVALID",active));}pane["activePane"]=json!(active);}map.insert("pane".into(),pane);},_=>{}}
 },Event::DocType(_)=>return Err(error("XML_DOCTYPE_FORBIDDEN",part)),Event::Eof=>break,_=>{}}buffer.clear();
 if merges.len() as u64*64+heights.len() as u64*48+hidden_rows.len() as u64*4>limits.max_metadata_bytes{return Err(error("XML_METADATA_BUDGET","Worksheet metadata exceeds budget"));}
 }
 map.insert("merges".into(),json!(merges));if !heights.is_empty(){map.insert("rowHeightsPx".into(),json!(heights));}if !widths.is_empty(){map.insert("nativeColumnWidths".into(),json!(widths));}if !hidden_rows.is_empty(){map.insert("nativeHiddenRows".into(),json!(hidden_rows));}if !hidden_columns.is_empty(){map.insert("hiddenColumns".into(),json!(hidden_columns));}Ok(map)
})
}
fn rewrite_sheet_directory(
    bytes: &[u8],
    sheets: &[SheetManifest],
    original: &[SheetMetadata],
    budget: u64,
) -> KernelResult<Vec<u8>> {
    let root = xmlnode::parse(bytes, budget)?;
    let parent = root
        .child("sheets")
        .ok_or_else(|| error("OOXML_SHEETS_MISSING", "sheets"))?;
    let mut out = Vec::new();
    let mut position = 0;
    for (node, sheet) in parent.children_named("sheet").zip(sheets) {
        let old = original.iter().find(|s| s.id == sheet.sheet_id).unwrap();
        let hidden = sheet
            .metadata
            .get("hidden")
            .and_then(Value::as_bool)
            .unwrap_or(false);
        if old.name == sheet.name && old.hidden == hidden {
            continue;
        }
        out.extend_from_slice(&bytes[position..node.span.start]);
        let mut tag = format!("<{}", node.name);
        for (k, v) in &node.attributes {
            if k != "name" && k != "state" {
                tag.push_str(&format!(" {k}=\"{}\"", xmlnode::escape(v)));
            }
        }
        tag.push_str(&format!(" name=\"{}\"", xmlnode::escape(&sheet.name)));
        if hidden {
            tag.push_str(
                if old.metadata.get("nativeVeryHidden") == Some(&json!(true)) {
                    " state=\"veryHidden\""
                } else {
                    " state=\"hidden\""
                },
            );
        }
        tag.push_str("/>");
        out.extend_from_slice(tag.as_bytes());
        position = node.span.end;
    }
    out.extend_from_slice(&bytes[position..]);
    Ok(out)
}
struct TemporaryParts {
    paths: Vec<PathBuf>,
}
impl TemporaryParts {
    fn new() -> Self {
        Self { paths: Vec::new() }
    }
    fn create(&mut self, directory: &Path, kind: &str) -> KernelResult<PathBuf> {
        static NEXT: AtomicU64 = AtomicU64::new(0);
        let n = NEXT.fetch_add(1, Ordering::Relaxed);
        let path = directory.join(format!("native-{}-{n}-{kind}.xml", std::process::id()));
        OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&path)
            .map_err(io_error)?;
        self.paths.push(path.clone());
        Ok(path)
    }
}
impl Drop for TemporaryParts {
    fn drop(&mut self) {
        for path in &self.paths {
            let _ = std::fs::remove_file(path);
        }
    }
}
fn io_error(e: impl std::fmt::Display) -> KernelError {
    error("NATIVE_IO_ERROR", e.to_string())
}
fn metadata_equal(a: Option<&Value>, b: Option<&Value>) -> bool {
    if a == b {
        return true;
    }
    let empty = |v: Option<&Value>| {
        v.is_none()
            || v.is_some_and(|v| {
                v.is_null()
                    || v.as_array().is_some_and(Vec::is_empty)
                    || v.as_object().is_some_and(serde_json::Map::is_empty)
                    || v == &json!({"kind":"none"})
            })
    };
    empty(a) && empty(b)
}
