//! Trusted native file boundary. Java authorizes identity and owns the database
//! transaction; this module only reads/writes revision-bound task-directory files.
use crate::KernelHost;
use kernel_core::*;
use kernel_native_document::{NativeDocument, ResourceLimits};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::{
    cell::RefCell,
    collections::BTreeMap,
    fs::{self, File, OpenOptions},
    io::{Read, Write},
    path::{Path, PathBuf},
};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct FilePage {
    descriptor: PageDescriptor,
    file_handle: String,
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct ImportRequest {
    unit_id: String,
    name: String,
    file_handle: String,
    output_directory: String,
    #[serde(default)]
    limits: ResourceLimits,
}
#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct ExportRequest {
    unit_id: String,
    revision: u64,
    format: String,
    file_handle: String,
    source_file_handle: Option<String>,
    source_revision: Option<u64>,
    source_checksum: Option<String>,
    pages_directory: String,
    #[serde(default)]
    limits: ResourceLimits,
}

pub(crate) fn dispatch(
    host: &mut KernelHost,
    operation: &str,
    params: Value,
) -> KernelResult<Value> {
    let root = task_root()?;
    match operation {
        "document.import" => {
            let request: ImportRequest = serde_json::from_value(params).map_err(protocol)?;
            import(&root, request)
        }
        "document.export" => {
            let request: ExportRequest =
                serde_json::from_value(params.clone()).map_err(protocol)?;
            let manifest = host.workbook(&params)?.manifest();
            export(&root, request, manifest)
        }
        _ => Err(KernelError::new("UNKNOWN_OPERATION", operation)),
    }
}
fn import(root: &Path, request: ImportRequest) -> KernelResult<Value> {
    if request.unit_id.is_empty() || request.name.is_empty() {
        return Err(KernelError::new(
            "WORKBOOK_INVALID",
            "Trusted unitId and name are required",
        ));
    }
    let source = existing(root, Path::new(&request.file_handle), false)?;
    let directory = existing(root, Path::new(&request.output_directory), true)?;
    if fs::read_dir(&directory).map_err(io_error)?.next().is_some() {
        return Err(KernelError::new(
            "TASK_DIRECTORY_NOT_EMPTY",
            "Import output directory must be empty",
        ));
    }
    let document = NativeDocument::open(&source, 0, request.limits.clone())?;
    let mut created = CreatedFiles::default();
    let mut entries = Vec::new();
    let mut builders: BTreeMap<PageKey, PageBuilder> = BTreeMap::new();
    let mut stripe: Option<(String, u32)> = None;
    document.visit_cells(&directory, &mut |record| {
        let key = PageKey::for_address(&record.address);
        let next = (key.sheet_id.clone(), key.page_row);
        if stripe.as_ref() != Some(&next) {
            flush(&mut builders, &directory, &mut entries, &mut created)?;
            stripe = Some(next);
        }
        if !builders.contains_key(&key) {
            if (builders.len() + 1) as u64 * 1024 * 1024 > request.limits.max_working_bytes {
                return Err(KernelError::new(
                    "PAGE_WORKING_BUDGET",
                    "Active native import page stripe exceeds its explicit memory budget",
                ));
            }
            builders.insert(key.clone(), PageBuilder::new(key.clone(), 0));
        }
        builders
            .get_mut(&key)
            .unwrap()
            .set_cell(&record.address, Some(&record.cell))
    })?;
    flush(&mut builders, &directory, &mut entries, &mut created)?;
    let sheets = document
        .sheets
        .iter()
        .map(|sheet| SheetManifest {
            sheet_id: sheet.id.clone(),
            name: sheet.name.clone(),
            row_count: MAX_ROWS,
            column_count: MAX_COLUMNS,
            metadata: sheet.metadata.clone(),
        })
        .collect();
    let mut manifest = WorkbookPages::create(request.unit_id, request.name, sheets)?.manifest();
    manifest.revision = 0;
    manifest.pages = entries.iter().map(|e| e.descriptor.clone()).collect();
    manifest.metadata = document.metadata.workbook_metadata.clone();
    manifest.metadata.insert(
        "nativeArtifact".into(),
        serde_json::to_value(&document.artifact.identity).map_err(protocol)?,
    );
    // Opening validates directory identities and page extents before publishing.
    WorkbookPages::open(manifest.clone())?;
    let pages_manifest = directory.join("pages.json");
    write_new(
        &pages_manifest,
        &serde_json::to_vec(&entries).map_err(protocol)?,
        &mut created,
    )?;
    let manifest_file = directory.join("manifest.json");
    write_new(
        &manifest_file,
        &serde_json::to_vec(&manifest).map_err(protocol)?,
        &mut created,
    )?;
    let total = created.paths.iter().try_fold(0u64, |sum, p| {
        fs::metadata(p)
            .map(|m| sum.saturating_add(m.len()))
            .map_err(io_error)
    })?;
    if total > request.limits.max_temporary_bytes {
        return Err(KernelError::new(
            "TEMPORARY_BYTES_LIMIT",
            "Native import output exceeds temporary storage budget",
        ));
    }
    created.committed = true;
    Ok(
        json!({"manifest":manifest,"outputDirectory":directory,"pagesManifestFile":pages_manifest,"artifact":document.artifact.identity,"metadata":document.metadata}),
    )
}
fn flush(
    builders: &mut BTreeMap<PageKey, PageBuilder>,
    directory: &Path,
    entries: &mut Vec<FilePage>,
    created: &mut CreatedFiles,
) -> KernelResult<()> {
    for (_, builder) in std::mem::take(builders) {
        let (descriptor, bytes) = builder.finish()?;
        let path = directory.join(format!("page-{:08}.bin", entries.len()));
        write_new(&path, &bytes, created)?;
        entries.push(FilePage {
            descriptor,
            file_handle: path.to_string_lossy().into_owned(),
        });
    }
    Ok(())
}
fn export(root: &Path, request: ExportRequest, manifest: WorkbookManifest) -> KernelResult<Value> {
    if manifest.unit_id != request.unit_id || manifest.revision != request.revision {
        return Err(KernelError::new(
            "REVISION_CONFLICT",
            "Export manifest identity is stale",
        ));
    }
    let directory = existing(root, Path::new(&request.pages_directory), true)?;
    let output = new_file(root, Path::new(&request.file_handle))?;
    let mut temporary_source = CreatedFiles::default();
    let document = match (
        &request.source_file_handle,
        request.source_revision,
        &request.source_checksum,
    ) {
        (Some(handle), Some(revision), Some(_)) => NativeDocument::open(
            existing(root, Path::new(handle), false)?,
            revision,
            request.limits.clone(),
        )?,
        (None, None, None) => {
            let path = directory.join("new-workbook-source.xlsx");
            let document = NativeDocument::create_source(&path, &manifest, request.limits.clone())?;
            temporary_source.paths.push(path);
            document
        }
        _ => {
            return Err(KernelError::new(
                "ARTIFACT_BINDING_REQUIRED",
                "Source file, revision and checksum must be supplied together",
            ))
        }
    };
    if request.format != document.format.as_str() {
        return Err(KernelError::new("UNSUPPORTED_FEATURE", "Export format conversion requires an explicit native conversion implementation").at(request.format));
    }
    let cells = FilePageReader::open(
        root,
        &directory,
        manifest.clone(),
        request.limits.max_working_bytes,
    )?;
    let source_revision = request
        .source_revision
        .unwrap_or(document.artifact.identity.revision);
    let source_checksum = request
        .source_checksum
        .as_deref()
        .unwrap_or(&document.artifact.identity.checksum);
    let artifact = document.export_to_path(
        &output,
        source_revision,
        source_checksum,
        &manifest,
        &cells,
        &directory,
    )?;
    Ok(
        json!({"unitId":request.unit_id,"revision":request.revision,"fileHandle":output,"artifact":artifact.identity,"metadata":document.metadata}),
    )
}
struct FilePageReader {
    manifest: WorkbookManifest,
    entries: BTreeMap<PageKey, FilePage>,
    cache: RefCell<WorkbookPages>,
}
impl FilePageReader {
    fn open(
        root: &Path,
        directory: &Path,
        manifest: WorkbookManifest,
        budget: u64,
    ) -> KernelResult<Self> {
        let path = existing(root, &directory.join("pages.json"), false)?;
        let file = File::open(path).map_err(io_error)?;
        let length = file.metadata().map_err(io_error)?.len();
        if length > 16 * 1024 * 1024 {
            return Err(KernelError::new(
                "PAGE_DIRECTORY_LIMIT",
                "Page file directory exceeds 16 MiB",
            ));
        }
        let entries: Vec<FilePage> = serde_json::from_reader(file).map_err(protocol)?;
        let mut index = BTreeMap::new();
        let expected: BTreeMap<_, _> = manifest.pages.iter().map(|p| (p.key(), p)).collect();
        for mut entry in entries {
            if expected.get(&entry.descriptor.key()).copied() != Some(&entry.descriptor) {
                return Err(KernelError::new(
                    "PAGE_DESCRIPTOR_MISMATCH",
                    "File page descriptor is not in the committed manifest",
                ));
            }
            let path = existing(root, Path::new(&entry.file_handle), false)?;
            if !path.starts_with(directory) {
                return Err(KernelError::new(
                    "TASK_PATH_ESCAPE",
                    "Export pages must remain in their task page directory",
                ));
            }
            entry.file_handle = path.to_string_lossy().into_owned();
            if index.insert(entry.descriptor.key(), entry).is_some() {
                return Err(KernelError::new(
                    "PAGE_DUPLICATE",
                    "Duplicate file page identity",
                ));
            }
        }
        if index.len() != manifest.pages.len() {
            return Err(KernelError::new(
                "DATA_PAGE_UNAVAILABLE",
                "Export page directory is incomplete",
            ));
        }
        let mut cache = WorkbookPages::open(manifest.clone())?;
        cache.set_page_budget(budget.min(64 * 1024 * 1024))?;
        Ok(Self {
            manifest,
            entries: index,
            cache: RefCell::new(cache),
        })
    }
    fn load(&self, key: &PageKey) -> KernelResult<()> {
        let Some(entry) = self.entries.get(key) else {
            return Ok(());
        };
        let mut cache = self.cache.borrow_mut();
        let probe = entry
            .descriptor
            .occupied_range
            .as_ref()
            .map(|r| CellAddress {
                sheet_id: key.sheet_id.clone(),
                row: r.start_row,
                column: r.start_column,
            })
            .ok_or_else(|| {
                KernelError::new(
                    "PAGE_STATS_INVALID",
                    "Stored nonempty page has no occupied range",
                )
            })?;
        match cache.read_cell(&probe) {
            Ok(_) => return Ok(()),
            Err(e) if e.code == "DATA_PAGE_UNAVAILABLE" => {}
            Err(e) => return Err(e),
        }
        let file = File::open(&entry.file_handle).map_err(io_error)?;
        let actual = file.metadata().map_err(io_error)?.len();
        if actual != entry.descriptor.byte_length as u64 || actual > 1024 * 1024 {
            return Err(KernelError::new(
                "PAGE_SIZE_LIMIT",
                "Stored page byte length differs from its manifest",
            ));
        }
        let mut bytes = Vec::with_capacity(actual as usize);
        file.take(actual + 1)
            .read_to_end(&mut bytes)
            .map_err(io_error)?;
        cache.load_page(&entry.descriptor, &bytes)
    }
}
impl CellReader for FilePageReader {
    fn revision(&self) -> u64 {
        self.manifest.revision
    }
    fn read_cell(&self, address: &CellAddress) -> KernelResult<Option<Cell>> {
        self.load(&PageKey::for_address(address))?;
        self.cache.borrow().read_cell(address)
    }
    fn read_range(
        &self,
        range: &RangeRef,
        visitor: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        range.validate()?;
        if !self
            .manifest
            .sheets
            .iter()
            .any(|s| s.sheet_id == range.sheet_id)
        {
            return Err(KernelError::new("SHEET_NOT_FOUND", &range.sheet_id));
        }
        let start = PageKey {
            sheet_id: range.sheet_id.clone(),
            page_row: range.start_row / PAGE_ROWS,
            page_column: 0,
        };
        let end = PageKey {
            sheet_id: range.sheet_id.clone(),
            page_row: range.end_row / PAGE_ROWS,
            page_column: MAX_COLUMNS / PAGE_COLUMNS,
        };
        for (key, entry) in self.entries.range(start..=end) {
            let Some(occupied) = &entry.descriptor.occupied_range else {
                continue;
            };
            if !occupied.intersects(range) {
                continue;
            }
            self.load(key)?;
            let clipped = RangeRef {
                sheet_id: range.sheet_id.clone(),
                start_row: range.start_row.max(occupied.start_row),
                end_row: range.end_row.min(occupied.end_row),
                start_column: range.start_column.max(occupied.start_column),
                end_column: range.end_column.min(occupied.end_column),
            };
            self.cache.borrow().read_range(&clipped, visitor)?;
        }
        Ok(())
    }
}
#[derive(Default)]
struct CreatedFiles {
    paths: Vec<PathBuf>,
    committed: bool,
}
impl Drop for CreatedFiles {
    fn drop(&mut self) {
        if !self.committed {
            for path in &self.paths {
                let _ = fs::remove_file(path);
            }
        }
    }
}
fn write_new(path: &Path, bytes: &[u8], created: &mut CreatedFiles) -> KernelResult<()> {
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(path)
        .map_err(io_error)?;
    created.paths.push(path.to_owned());
    file.write_all(bytes).map_err(io_error)?;
    file.sync_all().map_err(io_error)
}
fn task_root() -> KernelResult<PathBuf> {
    let value = std::env::var_os("KERNEL_TASK_DIRECTORY").ok_or_else(|| {
        KernelError::new(
            "NATIVE_TASK_ROOT_REQUIRED",
            "KERNEL_TASK_DIRECTORY must be explicitly configured",
        )
    })?;
    let root = fs::canonicalize(value).map_err(io_error)?;
    if !root.is_dir() {
        return Err(KernelError::new(
            "NATIVE_TASK_ROOT_INVALID",
            "Task root is not a directory",
        ));
    }
    Ok(root)
}
fn existing(root: &Path, path: &Path, directory: bool) -> KernelResult<PathBuf> {
    if !path.is_absolute() {
        return Err(KernelError::new(
            "TASK_PATH_INVALID",
            "File handles must be absolute native task paths",
        ));
    }
    let resolved = fs::canonicalize(path).map_err(io_error)?;
    if !resolved.starts_with(root) {
        return Err(KernelError::new(
            "TASK_PATH_ESCAPE",
            "File handle escapes the configured native task root",
        ));
    }
    if directory != resolved.is_dir() {
        return Err(KernelError::new(
            "TASK_PATH_TYPE_INVALID",
            "File handle has the wrong object type",
        ));
    }
    Ok(resolved)
}
fn new_file(root: &Path, path: &Path) -> KernelResult<PathBuf> {
    if path.exists() {
        return Err(KernelError::new(
            "ARTIFACT_OVERWRITE_FORBIDDEN",
            "Output artifact must not already exist",
        ));
    }
    let parent = path
        .parent()
        .ok_or_else(|| KernelError::new("TASK_PATH_INVALID", "Output parent is required"))?;
    let parent = existing(root, parent, true)?;
    let name = path
        .file_name()
        .ok_or_else(|| KernelError::new("TASK_PATH_INVALID", "Output filename is required"))?;
    Ok(parent.join(name))
}
fn io_error(e: impl std::fmt::Display) -> KernelError {
    KernelError::new("NATIVE_FILE_IO", e.to_string())
        .recover("Correct task-directory permissions or file handles and retry")
}
fn protocol(e: impl std::fmt::Display) -> KernelError {
    KernelError::new("KERNEL_PROTOCOL_ERROR", e.to_string())
}
