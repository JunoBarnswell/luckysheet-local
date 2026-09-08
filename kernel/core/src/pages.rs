//! Canonical columnar page store for workbook data.
//!
//! A page is 1024 rows by 32 columns.  The manifest is the authoritative
//! directory; an entry in the directory whose bytes have not been loaded is
//! deliberately distinguishable from a page which is not in the directory.

use crate::{
    AccessRole, Cell, CellAddress, CellReader, FormulaError, KernelError, KernelResult, MAX_COLUMNS, MAX_ROWS,
    PAGE_COLUMNS, PAGE_ROWS, RangeRef, Scalar, WORKBOOK_MANIFEST_VERSION,
};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::sync::Arc;

const PAGE_MAGIC: &[u8; 4] = b"LSPG";
const PAGE_ENCODING_VERSION: u8 = 1;
const DEFAULT_PAGE_BUDGET: u64 = 64 * 1024 * 1024;

fn default_workbook_metadata() -> BTreeMap<String, Value> {
    BTreeMap::from([
        ("dateSystem".into(), json!("1900")),
        ("numericContext".into(), json!({ "significantDigits": 15 })),
        (
            "collationContext".into(),
            json!({
                "cultureId": "invariant",
                "caseSensitive": true,
                "accentSensitive": true,
                "numericTextMode": "lexical",
                "blankOrder": "last",
                "typeOrder": ["number", "text", "boolean", "error", "blank"],
                "customLists": []
            }),
        ),
        (
            "calculationSettings".into(),
            json!({
                "mode": "automatic",
                "iterativeCalculation": false,
                "maximumIterations": 100,
                "maximumChange": 0.001,
                "precisionAsDisplayed": false,
                "calculateBeforeSave": true,
                "fullCalculationOnLoad": false
            }),
        ),
        (
            "dimensionMetrics".into(),
            json!({
                "normalFontFamily": "Calibri",
                "normalFontSizePx": 14.6666666667,
                "maximumDigitWidthPx": 7
            }),
        ),
        (
            "editingOptions".into(),
            json!({
                "allowEditDirectly": true,
                "moveAfterEnter": true,
                "enterDirection": "down",
                "formulaAutoComplete": true,
                "valueAutoComplete": true,
                "fixedDecimalPlaces": null
            }),
        ),
        ("definedNameModels".into(), json!([])),
        (
            "dataModel".into(),
            json!({
                "sources": [], "tables": [], "relationships": [], "views": []
            }),
        ),
    ])
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PageKey {
    pub sheet_id: String,
    pub page_row: u32,
    pub page_column: u32,
}

impl PageKey {
    pub fn for_address(address: &CellAddress) -> Self {
        Self {
            sheet_id: address.sheet_id.clone(),
            page_row: address.row / PAGE_ROWS,
            page_column: address.column / PAGE_COLUMNS,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SheetManifest {
    pub sheet_id: String,
    pub name: String,
    pub row_count: u32,
    pub column_count: u32,
    #[serde(default)]
    pub metadata: BTreeMap<String, Value>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PageDescriptor {
    pub sheet_id: String,
    pub page_row: u32,
    pub page_column: u32,
    pub revision: u64,
    pub checksum: String,
    pub byte_length: u32,
    pub cell_count: u32,
    pub occupied_range: Option<RangeRef>,
}
impl PageDescriptor {
    pub fn key(&self) -> PageKey {
        PageKey {
            sheet_id: self.sheet_id.clone(),
            page_row: self.page_row,
            page_column: self.page_column,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PagePayload {
    #[serde(flatten)]
    pub descriptor: PageDescriptor,
    pub payload_base64: String,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct WorkbookManifest {
    pub schema: String,
    pub version: u32,
    pub unit_id: String,
    pub name: String,
    pub revision: u64,
    pub sheets: Vec<SheetManifest>,
    pub pages: Vec<PageDescriptor>,
    #[serde(default)]
    pub metadata: BTreeMap<String, Value>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CellWrite {
    pub address: CellAddress,
    pub cell: Option<Cell>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ChangeSet {
    pub operation_id: String,
    pub base_revision: u64,
    pub revision: u64,
    pub manifest: WorkbookManifest,
    pub pages: Vec<PagePayload>,
    pub removed_pages: Vec<PageKey>,
    pub affected_ranges: Vec<RangeRef>,
    pub history: HistoryRecord,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct PageDelta {
    pub key: PageKey,
    pub before: Option<PageDescriptor>,
    pub after: Option<PageDescriptor>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ManifestMetadata {
    pub name: String,
    pub sheets: Vec<SheetManifest>,
    pub metadata: BTreeMap<String, Value>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct HistoryRecord {
    pub operation_id: String,
    pub base_revision: u64,
    pub revision: u64,
    pub required_role: AccessRole,
    pub page_deltas: Vec<PageDelta>,
    pub metadata_before: Option<ManifestMetadata>,
    pub metadata_after: Option<ManifestMetadata>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct SheetStats {
    pub sheet_id: String,
    pub cell_count: u64,
    pub occupied_range: Option<RangeRef>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Column {
    pub tags: Vec<u8>,
    pub numbers: Vec<f64>,
    pub text_ids: Vec<u32>,
    pub formula_ids: Vec<u32>,
    pub metadata_ids: Vec<u32>,
}
impl Default for Column {
    fn default() -> Self {
        Self {
            tags: vec![0; PAGE_ROWS as usize],
            numbers: vec![0.0; PAGE_ROWS as usize],
            text_ids: vec![0; PAGE_ROWS as usize],
            formula_ids: vec![0; PAGE_ROWS as usize],
            metadata_ids: vec![0; PAGE_ROWS as usize],
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CellPage {
    /// Exactly 32 slots, with storage allocated only for columns touched by a cell.
    pub columns: Vec<Option<Column>>,
    pub text_dictionary: Vec<String>,
    pub formula_dictionary: Vec<String>,
    pub metadata_dictionary: Vec<BTreeMap<String, Value>>,
    pub errors: Vec<FormulaError>,
    text_index: HashMap<String, u32>,
    formula_index: HashMap<String, u32>,
    error_index: HashMap<(String, String, String), u32>,
}
impl Default for CellPage {
    fn default() -> Self {
        Self {
            columns: (0..PAGE_COLUMNS).map(|_| None).collect(),
            text_dictionary: Vec::new(),
            formula_dictionary: Vec::new(),
            metadata_dictionary: Vec::new(),
            errors: Vec::new(),
            text_index: HashMap::new(),
            formula_index: HashMap::new(),
            error_index: HashMap::new(),
        }
    }
}

pub struct PageBuilder {
    key: PageKey,
    revision: u64,
    page: CellPage,
}

impl PageBuilder {
    pub fn new(key: PageKey, revision: u64) -> Self {
        Self {
            key,
            revision,
            page: CellPage::default(),
        }
    }

    pub fn set_cell(&mut self, address: &CellAddress, cell: Option<&Cell>) -> KernelResult<()> {
        if PageKey::for_address(address) != self.key {
            return Err(KernelError::new(
                "CELL_ADDRESS_INVALID",
                "Cell is outside the page builder key",
            ));
        }
        address.validate()?;
        set_cell(&mut self.page, address, cell)
    }

    pub fn finish(self) -> KernelResult<(PageDescriptor, Vec<u8>)> {
        let bytes = encode_page(&self.page)?;
        if bytes.len() > 1024 * 1024 {
            return Err(KernelError::new(
                "PAGE_SIZE_LIMIT",
                "Encoded page exceeds 1 MiB",
            ));
        }
        let (cell_count, occupied_range) = page_stats(&self.page, &self.key);
        Ok((
            PageDescriptor {
                sheet_id: self.key.sheet_id,
                page_row: self.key.page_row,
                page_column: self.key.page_column,
                revision: self.revision,
                checksum: checksum(&bytes),
                byte_length: bytes.len() as u32,
                cell_count,
                occupied_range,
            },
            bytes,
        ))
    }
}

#[derive(Clone)]
struct PageStore {
    page: Arc<CellPage>,
    dirty: bool,
    bytes: u64,
}

#[derive(Clone)]
pub struct WorkbookPages {
    manifest: WorkbookManifest,
    pages: BTreeMap<PageKey, PageStore>,
    unavailable: BTreeSet<PageKey>,
    page_budget_bytes: u64,
    loaded_bytes: u64,
}

impl WorkbookPages {
    pub fn open_reusing(
        manifest: WorkbookManifest,
        previous: &WorkbookPages,
    ) -> KernelResult<Self> {
        let mut next = Self::open(manifest)?;
        for descriptor in &next.manifest.pages {
            let key = descriptor.key();
            if let Some(old) = previous.pages.get(&key) {
                if previous.descriptor(&key) == Some(descriptor) {
                    next.loaded_bytes += old.bytes;
                    next.pages.insert(key.clone(), old.clone());
                    next.unavailable.remove(&key);
                }
            }
        }
        if next.loaded_bytes > next.page_budget_bytes {
            next.evict_for(0)?;
        }
        Ok(next)
    }
    pub fn create(
        unit_id: impl Into<String>,
        name: impl Into<String>,
        sheets: Vec<SheetManifest>,
    ) -> KernelResult<Self> {
        validate_sheets(&sheets)?;
        let unit_id = unit_id.into();
        let name = name.into();
        if unit_id.is_empty() {
            return Err(KernelError::new("WORKBOOK_INVALID", "unitId is required"));
        }
        Ok(Self {
            manifest: WorkbookManifest {
                schema: "WorkbookManifest".into(),
                version: WORKBOOK_MANIFEST_VERSION,
                unit_id,
                name,
                revision: 0,
                sheets,
                pages: Vec::new(),
                metadata: default_workbook_metadata(),
            },
            pages: BTreeMap::new(),
            unavailable: BTreeSet::new(),
            page_budget_bytes: DEFAULT_PAGE_BUDGET,
            loaded_bytes: 0,
        })
    }

    pub fn open(manifest: WorkbookManifest) -> KernelResult<Self> {
        validate_manifest(&manifest)?;
        let unavailable = manifest.pages.iter().map(PageDescriptor::key).collect();
        Ok(Self {
            manifest,
            pages: BTreeMap::new(),
            unavailable,
            page_budget_bytes: DEFAULT_PAGE_BUDGET,
            loaded_bytes: 0,
        })
    }
    pub fn manifest(&self) -> WorkbookManifest {
        self.manifest.clone()
    }

    /// Publish a historical manifest as a new revision using immutable page
    /// references. The persistence owner must prove any referenced page bytes
    /// before committing this transition.
    pub fn restore_manifest(
        &mut self,
        operation_id: String,
        base_revision: u64,
        mut target: WorkbookManifest,
    ) -> KernelResult<ChangeSet> {
        if base_revision != self.manifest.revision || target.revision > base_revision {
            return Err(KernelError::new(
                "STALE_REVISION",
                "Restore revision is not based on the current workbook",
            ));
        }
        if target.unit_id != self.manifest.unit_id || operation_id.is_empty() {
            return Err(KernelError::new(
                "HISTORY_INVALID",
                "Restore requires the same workbook and an operation identity",
            ));
        }
        validate_manifest(&target)?;
        target.revision = base_revision
            .checked_add(1)
            .ok_or_else(|| KernelError::new("REVISION_OVERFLOW", "Workbook revision exhausted"))?;
        let before: BTreeMap<_, _> = self
            .manifest
            .pages
            .iter()
            .map(|p| (p.key(), p.clone()))
            .collect();
        let after: BTreeMap<_, _> = target.pages.iter().map(|p| (p.key(), p.clone())).collect();
        let keys: BTreeSet<_> = before.keys().chain(after.keys()).cloned().collect();
        let page_deltas: Vec<_> = keys
            .into_iter()
            .filter(|k| before.get(k) != after.get(k))
            .map(|key| PageDelta {
                before: before.get(&key).cloned(),
                after: after.get(&key).cloned(),
                key,
            })
            .collect();
        let metadata_changed = self.manifest.name != target.name
            || self.manifest.sheets != target.sheets
            || self.manifest.metadata != target.metadata;
        let history = HistoryRecord {
            operation_id: operation_id.clone(),
            base_revision,
            revision: target.revision,
            required_role: AccessRole::Editor,
            page_deltas: page_deltas.clone(),
            metadata_before: metadata_changed.then(|| ManifestMetadata {
                name: self.manifest.name.clone(),
                sheets: self.manifest.sheets.clone(),
                metadata: self.manifest.metadata.clone(),
            }),
            metadata_after: metadata_changed.then(|| ManifestMetadata {
                name: target.name.clone(),
                sheets: target.sheets.clone(),
                metadata: target.metadata.clone(),
            }),
        };
        let mut affected_ranges = Vec::new();
        if metadata_changed {
            for sheet in self.manifest.sheets.iter().chain(target.sheets.iter()) {
                let range = RangeRef {
                    sheet_id: sheet.sheet_id.clone(),
                    start_row: 0,
                    end_row: sheet.row_count - 1,
                    start_column: 0,
                    end_column: sheet.column_count - 1,
                };
                if !affected_ranges.contains(&range) {
                    affected_ranges.push(range);
                }
            }
        } else {
            affected_ranges.extend(page_deltas.iter().filter_map(|delta| {
                delta
                    .after
                    .as_ref()
                    .and_then(|d| d.occupied_range.clone())
                    .or_else(|| delta.before.as_ref().and_then(|d| d.occupied_range.clone()))
            }));
        }
        let next = Self::open_reusing(target.clone(), self)?;
        let changes = ChangeSet {
            operation_id,
            base_revision,
            revision: target.revision,
            manifest: target,
            pages: Vec::new(),
            removed_pages: page_deltas
                .iter()
                .filter(|p| p.after.is_none())
                .map(|p| p.key.clone())
                .collect(),
            affected_ranges,
            history,
        };
        *self = next;
        Ok(changes)
    }

    /// Reverts one committed operation only when every owned page and metadata
    /// value still equals that operation's after state. Later overlapping work
    /// therefore rejects atomically instead of being overwritten.
    pub fn undo_history(
        &mut self,
        operation_id: String,
        base_revision: u64,
        record: HistoryRecord,
    ) -> KernelResult<ChangeSet> {
        if base_revision != self.manifest.revision {
            return Err(
                KernelError::new("STALE_REVISION", "Undo base revision is stale")
                    .recover("refresh-manifest"),
            );
        }
        if record.operation_id.trim().is_empty()
            || record.revision
                != record.base_revision.checked_add(1).ok_or_else(|| {
                    KernelError::new("HISTORY_INVALID", "History revision overflow")
                })?
            || record.revision > base_revision
            || record.metadata_before.is_some() != record.metadata_after.is_some()
        {
            return Err(KernelError::new(
                "HISTORY_INVALID",
                "Undo history has an invalid revision or metadata transition",
            ));
        }
        let current: BTreeMap<_, _> = self
            .manifest
            .pages
            .iter()
            .cloned()
            .map(|d| (d.key(), d))
            .collect();
        let mut seen = BTreeSet::new();
        for delta in &record.page_deltas {
            if !seen.insert(delta.key.clone())
                || delta.before == delta.after
                || current.get(&delta.key) != delta.after.as_ref()
            {
                return Err(KernelError::new(
                    "UNDO_CONFLICT",
                    "A later operation changed a page owned by the undo target",
                )
                .at(format!(
                    "{}:{}:{}",
                    delta.key.sheet_id, delta.key.page_row, delta.key.page_column
                ))
                .recover("refresh-history"));
            }
        }
        let mut target = self.manifest.clone();
        if let (Some(before), Some(after)) = (&record.metadata_before, &record.metadata_after) {
            let current_metadata = ManifestMetadata {
                name: target.name.clone(),
                sheets: target.sheets.clone(),
                metadata: target.metadata.clone(),
            };
            if &current_metadata != after {
                return Err(KernelError::new(
                    "UNDO_CONFLICT",
                    "A later operation changed workbook metadata owned by the undo target",
                )
                .recover("refresh-history"));
            }
            target.name = before.name.clone();
            target.sheets = before.sheets.clone();
            target.metadata = before.metadata.clone();
        }
        target
            .pages
            .retain(|descriptor| !seen.contains(&descriptor.key()));
        for delta in &record.page_deltas {
            if let Some(before) = &delta.before {
                target.pages.push(before.clone());
            }
        }
        target.pages.sort_by_key(PageDescriptor::key);
        self.restore_manifest(operation_id, base_revision, target)
    }
    pub fn set_page_budget(&mut self, bytes: u64) -> KernelResult<()> {
        if bytes == 0 {
            return Err(KernelError::new(
                "PAGE_BUDGET_INVALID",
                "Page budget must be positive",
            ));
        }
        self.page_budget_bytes = bytes;
        self.evict_clean();
        if self.loaded_bytes > bytes {
            return Err(KernelError::new(
                "PAGE_BUDGET_EXCEEDED",
                "Loaded pages exceed page budget",
            ));
        }
        Ok(())
    }
    pub fn load_page(&mut self, descriptor: &PageDescriptor, bytes: &[u8]) -> KernelResult<()> {
        let expected = self.descriptor(&descriptor.key()).ok_or_else(|| {
            KernelError::new("PAGE_UNKNOWN", "Page is not present in the manifest")
        })?;
        if expected != descriptor {
            return Err(KernelError::new(
                "PAGE_DESCRIPTOR_MISMATCH",
                "Page descriptor differs from manifest",
            ));
        }
        if bytes.len() != descriptor.byte_length as usize {
            return Err(KernelError::new(
                "PAGE_LENGTH_INVALID",
                "Page byte length differs from descriptor",
            ));
        }
        if checksum(bytes) != descriptor.checksum {
            return Err(KernelError::new(
                "PAGE_CHECKSUM_INVALID",
                "Page checksum does not match descriptor",
            )
            .recover("reload-page"));
        }
        let page = decode_page(bytes)?;
        let (actual_count, actual_range) = page_stats(&page, &descriptor.key());
        if actual_count != descriptor.cell_count || actual_range != descriptor.occupied_range {
            return Err(KernelError::new(
                "PAGE_STATS_INVALID",
                "Page statistics do not match encoded cells",
            ));
        }
        let size = bytes.len() as u64;
        if size > self.page_budget_bytes {
            return Err(KernelError::new(
                "PAGE_BUDGET_EXCEEDED",
                "Page exceeds configured page budget",
            ));
        }
        let old_size = self
            .pages
            .get(&descriptor.key())
            .map(|s| s.bytes)
            .unwrap_or(0);
        self.evict_for(size.saturating_sub(old_size))?;
        if let Some(old) = self.pages.remove(&descriptor.key()) {
            self.loaded_bytes -= old.bytes;
        }
        self.loaded_bytes += size;
        self.pages.insert(
            descriptor.key(),
            PageStore {
                page: Arc::new(page),
                dirty: false,
                bytes: size,
            },
        );
        self.unavailable.remove(&descriptor.key());
        Ok(())
    }
    pub fn page_bytes(&self, key: &PageKey) -> KernelResult<Vec<u8>> {
        let store = self.pages.get(key).ok_or_else(|| {
            if self.unavailable.contains(key) {
                KernelError::new("DATA_PAGE_UNAVAILABLE", "Page bytes have not been loaded").at(
                    format!("{}:{}:{}", key.sheet_id, key.page_row, key.page_column),
                )
            } else {
                KernelError::new("PAGE_NOT_FOUND", "Page is not present in the manifest")
            }
        })?;
        encode_page(&store.page)
    }

    pub fn sheet_stats(&self, sheet_id: &str) -> KernelResult<SheetStats> {
        let sheet = self.sheet_id(sheet_id)?;
        let mut cell_count = 0u64;
        let mut occupied_range: Option<RangeRef> = None;
        for descriptor in self
            .manifest
            .pages
            .iter()
            .filter(|d| d.sheet_id == sheet_id)
        {
            cell_count = cell_count.saturating_add(descriptor.cell_count as u64);
            if let Some(range) = &descriptor.occupied_range {
                occupied_range = Some(match occupied_range {
                    None => range.clone(),
                    Some(mut current) => {
                        current.start_row = current.start_row.min(range.start_row);
                        current.end_row = current.end_row.max(range.end_row);
                        current.start_column = current.start_column.min(range.start_column);
                        current.end_column = current.end_column.max(range.end_column);
                        current
                    }
                });
            }
        }
        if cell_count > (sheet.row_count as u64) * (sheet.column_count as u64) {
            return Err(KernelError::new(
                "MANIFEST_INVALID",
                "Directory cell count exceeds worksheet capacity",
            ));
        }
        Ok(SheetStats {
            sheet_id: sheet_id.into(),
            cell_count,
            occupied_range,
        })
    }

    /// Resolves Excel's current-region boundary inside the canonical page
    /// store. The entire calculation stays in Rust so a UI projection never
    /// performs one host call per cell.
    pub fn resolve_current_region(
        &self,
        sheet_id: &str,
        active_row: u32,
        active_column: u32,
    ) -> KernelResult<RangeRef> {
        let sheet = self.sheet_id(sheet_id)?;
        if active_row >= sheet.row_count || active_column >= sheet.column_count {
            return Err(KernelError::new(
                "CELL_ADDRESS_INVALID",
                "Current-region address is outside the worksheet",
            ));
        }
        let single = || RangeRef {
            sheet_id: sheet_id.into(),
            start_row: active_row,
            end_row: active_row,
            start_column: active_column,
            end_column: active_column,
        };
        let Some(occupied_range) = self.sheet_stats(sheet_id)?.occupied_range else {
            return Ok(single());
        };
        let mut occupied = HashSet::new();
        self.read_range(&occupied_range, &mut |address, cell| {
            if cell.formula.is_some() || !matches!(cell.value, Scalar::Null) {
                occupied.insert((address.row, address.column));
            }
            Ok(())
        })?;
        if !occupied.contains(&(active_row, active_column)) {
            return Ok(single());
        }
        let mut start_row = active_row;
        let mut end_row = active_row;
        let mut start_column = active_column;
        let mut end_column = active_column;
        let mut grew = true;
        while grew {
            grew = false;
            while start_row > 0
                && (start_column..=end_column)
                    .any(|column| occupied.contains(&(start_row - 1, column)))
            {
                start_row -= 1;
                grew = true;
            }
            while end_row + 1 < sheet.row_count
                && (start_column..=end_column)
                    .any(|column| occupied.contains(&(end_row + 1, column)))
            {
                end_row += 1;
                grew = true;
            }
            while start_column > 0
                && (start_row..=end_row).any(|row| occupied.contains(&(row, start_column - 1)))
            {
                start_column -= 1;
                grew = true;
            }
            while end_column + 1 < sheet.column_count
                && (start_row..=end_row).any(|row| occupied.contains(&(row, end_column + 1)))
            {
                end_column += 1;
                grew = true;
            }
        }
        Ok(RangeRef {
            sheet_id: sheet_id.into(),
            start_row,
            end_row,
            start_column,
            end_column,
        })
    }

    pub fn apply_writes(
        &mut self,
        operation_id: impl Into<String>,
        base_revision: u64,
        writes: Vec<CellWrite>,
    ) -> KernelResult<ChangeSet> {
        self.apply_transaction(operation_id, base_revision, writes, self.manifest.clone())
    }

    /// Applies sheet/workbook metadata and cell writes as one revision.
    /// `next_manifest` supplies the caller's fully validated metadata view; its
    /// directory and revision are replaced by the transaction result.
    pub fn apply_transaction(
        &mut self,
        operation_id: impl Into<String>,
        base_revision: u64,
        writes: Vec<CellWrite>,
        next_manifest: WorkbookManifest,
    ) -> KernelResult<ChangeSet> {
        let operation_id = operation_id.into();
        if operation_id.trim().is_empty() {
            return Err(KernelError::new(
                "OPERATION_INVALID",
                "operationId is required",
            ));
        }
        if base_revision != self.manifest.revision {
            return Err(KernelError::new("STALE_REVISION", "Base revision is stale")
                .recover("refresh-manifest"));
        }
        if next_manifest.schema != "WorkbookManifest"
            || next_manifest.version != WORKBOOK_MANIFEST_VERSION
            || next_manifest.unit_id != self.manifest.unit_id
            || next_manifest.revision != base_revision
        {
            return Err(KernelError::new(
                "MANIFEST_INVALID",
                "Transaction manifest must identify the current workbook version",
            ));
        }
        validate_sheets(&next_manifest.sheets)?;
        let next_revision = base_revision
            .checked_add(1)
            .ok_or_else(|| KernelError::new("REVISION_OVERFLOW", "Workbook revision exhausted"))?;
        let mut touched: BTreeMap<PageKey, CellPage> = BTreeMap::new();
        let next_sheets: BTreeMap<&str, &SheetManifest> = next_manifest
            .sheets
            .iter()
            .map(|s| (s.sheet_id.as_str(), s))
            .collect();
        for write in &writes {
            write.address.validate()?;
            if let Some(cell) = &write.cell {
                cell.value.validate()?;
            }
            let target = next_sheets.get(write.address.sheet_id.as_str());
            let inside = target.is_some_and(|s| {
                write.address.row < s.row_count && write.address.column < s.column_count
            });
            // Clearing an address that was removed by this same structural
            // transaction is valid. A value write beyond the new extent is not.
            if !inside {
                let existed = self.manifest.sheets.iter().any(|s| {
                    s.sheet_id == write.address.sheet_id
                        && write.address.row < s.row_count
                        && write.address.column < s.column_count
                });
                if write.cell.is_some() || !existed {
                    return Err(KernelError::new(
                        "CELL_ADDRESS_INVALID",
                        "Write is outside the transaction worksheet extent",
                    ));
                }
                if target.is_none() {
                    continue;
                }
            }
            let key = PageKey::for_address(&write.address);
            if self.unavailable.contains(&key) {
                return Err(KernelError::new(
                    "DATA_PAGE_UNAVAILABLE",
                    "Cannot write an unloaded page",
                )
                .at(format!(
                    "{}:{}:{}",
                    key.sheet_id, key.page_row, key.page_column
                )));
            }
            let page = touched.entry(key.clone()).or_insert_with(|| {
                self.pages
                    .get(&key)
                    .map(|p| (*p.page).clone())
                    .unwrap_or_default()
            });
            set_cell(page, &write.address, write.cell.as_ref())?;
        }
        let mut directory: BTreeMap<PageKey, PageDescriptor> = self
            .manifest
            .pages
            .iter()
            .cloned()
            .map(|d| (d.key(), d))
            .collect();
        let mut removed = BTreeSet::new();
        // Extent changes own removal of cells beyond the final logical extent.
        // Boundary pages are rewritten in the same transaction; no intermediate
        // clear revision is ever exposed.
        for (key, descriptor) in &directory {
            let Some(sheet) = next_sheets.get(key.sheet_id.as_str()) else {
                removed.insert(key.clone());
                continue;
            };
            let row_start = key.page_row * PAGE_ROWS;
            let column_start = key.page_column * PAGE_COLUMNS;
            if row_start >= sheet.row_count || column_start >= sheet.column_count {
                removed.insert(key.clone());
                continue;
            }
            if descriptor
                .occupied_range
                .as_ref()
                .is_some_and(|r| r.end_row >= sheet.row_count || r.end_column >= sheet.column_count)
            {
                if self.unavailable.contains(key) {
                    return Err(KernelError::new(
                        "DATA_PAGE_UNAVAILABLE",
                        "Structural change requires the boundary page",
                    )
                    .at(format!(
                        "{}:{}:{}",
                        key.sheet_id, key.page_row, key.page_column
                    )));
                }
                let page = touched.entry(key.clone()).or_insert_with(|| {
                    (*self.pages.get(key).expect("manifest page is resident").page).clone()
                });
                for column in 0..PAGE_COLUMNS {
                    for row in 0..PAGE_ROWS {
                        if row_start + row >= sheet.row_count
                            || column_start + column >= sheet.column_count
                        {
                            set_cell(
                                page,
                                &CellAddress {
                                    sheet_id: key.sheet_id.clone(),
                                    row: row_start + row,
                                    column: column_start + column,
                                },
                                None,
                            )?;
                        }
                    }
                }
            }
        }
        let mut replacements: BTreeMap<PageKey, PageStore> = BTreeMap::new();
        let mut changed = Vec::new();
        for (key, page) in touched {
            if removed.contains(&key) {
                continue;
            }
            let (cell_count, occupied_range) = page_stats(&page, &key);
            if cell_count == 0 {
                if directory.contains_key(&key) {
                    removed.insert(key);
                }
                continue;
            }
            let bytes = encode_page(&page)?;
            if bytes.len() > 1024 * 1024 {
                return Err(KernelError::new(
                    "PAGE_SIZE_LIMIT",
                    "Encoded page exceeds 1 MiB",
                ));
            }
            let digest = checksum(&bytes);
            if directory.get(&key).is_some_and(|old| {
                old.checksum == digest
                    && old.cell_count == cell_count
                    && old.occupied_range == occupied_range
            }) {
                continue;
            }
            let descriptor = PageDescriptor {
                sheet_id: key.sheet_id.clone(),
                page_row: key.page_row,
                page_column: key.page_column,
                revision: next_revision,
                checksum: digest,
                byte_length: bytes.len() as u32,
                cell_count,
                occupied_range,
            };
            directory.insert(key.clone(), descriptor.clone());
            changed.push(PagePayload {
                descriptor,
                payload_base64: b64_encode(&bytes),
            });
            replacements.insert(
                key,
                PageStore {
                    page: Arc::new(page),
                    dirty: true,
                    bytes: bytes.len() as u64,
                },
            );
        }
        for key in &removed {
            directory.remove(key);
        }
        let mut committed = next_manifest;
        committed.revision = next_revision;
        committed.pages = directory.into_values().collect();
        validate_manifest(&committed)?;
        let next_descriptors: BTreeMap<PageKey, PageDescriptor> = committed
            .pages
            .iter()
            .cloned()
            .map(|d| (d.key(), d))
            .collect();
        let previous_descriptors: BTreeMap<PageKey, PageDescriptor> = self
            .manifest
            .pages
            .iter()
            .cloned()
            .map(|d| (d.key(), d))
            .collect();
        let all_keys: BTreeSet<PageKey> = previous_descriptors
            .keys()
            .chain(next_descriptors.keys())
            .cloned()
            .collect();
        let page_deltas: Vec<PageDelta> = all_keys
            .into_iter()
            .filter_map(|key| {
                let before = previous_descriptors.get(&key).cloned();
                let after = next_descriptors.get(&key).cloned();
                if before == after {
                    None
                } else {
                    Some(PageDelta { key, before, after })
                }
            })
            .collect();
        let metadata_before = if committed.name != self.manifest.name
            || committed.sheets != self.manifest.sheets
            || committed.metadata != self.manifest.metadata
        {
            Some(ManifestMetadata {
                name: self.manifest.name.clone(),
                sheets: self.manifest.sheets.clone(),
                metadata: self.manifest.metadata.clone(),
            })
        } else {
            None
        };
        let metadata_after = metadata_before.as_ref().map(|_| ManifestMetadata {
            name: committed.name.clone(),
            sheets: committed.sheets.clone(),
            metadata: committed.metadata.clone(),
        });
        // Reserve and evict against a private Arc-backed candidate. Every error
        // leaves both the committed directory and its resident cache untouched.
        let mut candidate = self.clone();
        for key in &removed {
            if let Some(page) = candidate.pages.remove(key) {
                candidate.loaded_bytes -= page.bytes;
            }
            candidate.unavailable.remove(key);
        }
        for key in replacements.keys() {
            if let Some(page) = candidate.pages.remove(key) {
                candidate.loaded_bytes -= page.bytes;
            }
        }
        let needed: u64 = replacements.values().map(|p| p.bytes).sum();
        candidate.evict_for(needed)?;
        for (key, page) in replacements {
            candidate.loaded_bytes += page.bytes;
            candidate.unavailable.remove(&key);
            candidate.pages.insert(key, page);
        }
        candidate.manifest = committed;
        // Dirty scope is page-bounded, not a million-element address list.
        let affected_ranges = page_deltas
            .iter()
            .filter_map(|delta| {
                delta
                    .after
                    .as_ref()
                    .and_then(|d| d.occupied_range.clone())
                    .or_else(|| delta.before.as_ref().and_then(|d| d.occupied_range.clone()))
            })
            .collect();
        let history = HistoryRecord {
            operation_id: operation_id.clone(),
            base_revision,
            revision: next_revision,
            required_role: AccessRole::Editor,
            page_deltas,
            metadata_before,
            metadata_after,
        };
        let result = ChangeSet {
            operation_id,
            base_revision,
            revision: next_revision,
            manifest: candidate.manifest.clone(),
            pages: changed,
            removed_pages: removed.into_iter().collect(),
            affected_ranges,
            history,
        };
        *self = candidate;
        Ok(result)
    }

    /// Restores descriptor-referenced historical pages supplied by the host.
    /// The host must prove each byte payload against `HistoryRecord.page_deltas`;
    /// no inverse cell operation is accepted as an undo substitute.
    pub fn apply_history(
        &mut self,
        record: HistoryRecord,
        old_pages: Vec<PagePayload>,
    ) -> KernelResult<WorkbookManifest> {
        if self.manifest.revision != record.revision {
            return Err(KernelError::new(
                "STALE_REVISION",
                "History record is not based on the current revision",
            )
            .recover("refresh-history"));
        }
        let mut payloads = BTreeMap::new();
        for payload in old_pages {
            let key = payload.descriptor.key();
            let delta = record
                .page_deltas
                .iter()
                .find(|d| d.key == key)
                .ok_or_else(|| {
                    KernelError::new(
                        "HISTORY_INVALID",
                        "Historical page is not referenced by the record",
                    )
                })?;
            if delta.before.as_ref() != Some(&payload.descriptor) {
                return Err(KernelError::new(
                    "HISTORY_INVALID",
                    "Historical page descriptor does not match before state",
                ));
            }
            let bytes = b64_decode(&payload.payload_base64)?;
            if bytes.len() != payload.descriptor.byte_length as usize
                || checksum(&bytes) != payload.descriptor.checksum
            {
                return Err(KernelError::new(
                    "PAGE_CHECKSUM_INVALID",
                    "Historical page proof failed",
                ));
            }
            let page = decode_page(&bytes)?;
            let (count, range) = page_stats(&page, &key);
            if count != payload.descriptor.cell_count || range != payload.descriptor.occupied_range
            {
                return Err(KernelError::new(
                    "PAGE_STATS_INVALID",
                    "Historical page statistics do not match",
                ));
            }
            payloads.insert(key, (payload.descriptor, page, bytes.len() as u64));
        }
        for delta in &record.page_deltas {
            if delta.before.is_some() && !payloads.contains_key(&delta.key) {
                return Err(KernelError::new(
                    "DATA_PAGE_UNAVAILABLE",
                    "Historical page bytes are required for undo",
                ));
            }
        }
        let mut target = self.manifest.clone();
        if let Some(meta) = &record.metadata_before {
            target.name = meta.name.clone();
            target.sheets = meta.sheets.clone();
            target.metadata = meta.metadata.clone();
        }
        target
            .pages
            .retain(|d| !record.page_deltas.iter().any(|delta| delta.key == d.key()));
        for delta in &record.page_deltas {
            if let Some(before) = &delta.before {
                target.pages.push(before.clone());
            }
        }
        target.pages.sort_by_key(PageDescriptor::key);
        target.revision = record.base_revision;
        validate_manifest(&target)?;
        let needed: u64 = payloads.values().map(|(_, _, size)| *size).sum();
        self.evict_for(needed)?;
        for delta in &record.page_deltas {
            if let Some(old) = self.pages.remove(&delta.key) {
                self.loaded_bytes -= old.bytes;
            }
            self.unavailable.remove(&delta.key);
            if let Some((_, page, size)) = payloads.remove(&delta.key) {
                self.loaded_bytes += size;
                self.pages.insert(
                    delta.key.clone(),
                    PageStore {
                        page: Arc::new(page),
                        dirty: true,
                        bytes: size,
                    },
                );
            }
        }
        self.manifest = target;
        Ok(self.manifest.clone())
    }

    fn descriptor(&self, key: &PageKey) -> Option<&PageDescriptor> {
        self.manifest.pages.iter().find(|d| d.key() == *key)
    }
    fn evict_for(&mut self, needed: u64) -> KernelResult<()> {
        while self.loaded_bytes.saturating_add(needed) > self.page_budget_bytes {
            let key = self
                .pages
                .iter()
                .find(|(_, s)| !s.dirty)
                .map(|(k, _)| k.clone());
            let Some(key) = key else {
                return Err(KernelError::new(
                    "PAGE_BUDGET_EXCEEDED",
                    "Loaded dirty pages exceed page budget",
                ));
            };
            if let Some(s) = self.pages.remove(&key) {
                self.loaded_bytes -= s.bytes;
                self.unavailable.insert(key);
            }
        }
        Ok(())
    }
    fn evict_clean(&mut self) {
        let _ = self.evict_for(0);
    }
}

impl CellReader for WorkbookPages {
    fn revision(&self) -> u64 {
        self.manifest.revision
    }
    fn read_cell(&self, address: &CellAddress) -> KernelResult<Option<Cell>> {
        address.validate()?;
        self.sheet_id(&address.sheet_id)?;
        let key = PageKey::for_address(address);
        match self.pages.get(&key) {
            Some(s) => get_cell(&s.page, address),
            None if self.unavailable.contains(&key) => Err(KernelError::new(
                "DATA_PAGE_UNAVAILABLE",
                "Page bytes have not been loaded",
            )
            .at(format!(
                "{}:{}:{}",
                key.sheet_id, key.page_row, key.page_column
            ))),
            None => Ok(None),
        }
    }
    fn read_range(
        &self,
        range: &RangeRef,
        visitor: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        range.validate()?;
        self.sheet_id(&range.sheet_id)?;
        let mut groups: BTreeMap<u32, Vec<(&PageDescriptor, &CellPage)>> = BTreeMap::new();
        for descriptor in &self.manifest.pages {
            if descriptor.sheet_id != range.sheet_id
                || !descriptor
                    .occupied_range
                    .as_ref()
                    .is_some_and(|r| r.intersects(range))
            {
                continue;
            }
            let key = descriptor.key();
            let page = self.pages.get(&key).ok_or_else(|| {
                KernelError::new("DATA_PAGE_UNAVAILABLE", "Range requires an unloaded page").at(
                    format!("{}:{}:{}", key.sheet_id, key.page_row, key.page_column),
                )
            })?;
            groups
                .entry(key.page_row)
                .or_default()
                .push((descriptor, &page.page));
        }
        // Page directories and column tags skip implicit blank space. Within a
        // page strip the visitor remains row-major for deterministic formulas.
        for (page_row, mut pages) in groups {
            pages.sort_by_key(|(descriptor, _)| descriptor.page_column);
            let first_row = range.start_row.max(page_row * PAGE_ROWS);
            let last_row = range.end_row.min((page_row + 1) * PAGE_ROWS - 1);
            for row in first_row..=last_row {
                for (descriptor, page) in &pages {
                    let occupied = descriptor
                        .occupied_range
                        .as_ref()
                        .expect("directory statistics are validated");
                    if row < occupied.start_row || row > occupied.end_row {
                        continue;
                    }
                    let first_column = range.start_column.max(occupied.start_column);
                    let last_column = range.end_column.min(occupied.end_column);
                    for column in first_column..=last_column {
                        let slot = (row % PAGE_ROWS) as usize;
                        if page.columns[(column % PAGE_COLUMNS) as usize]
                            .as_ref()
                            .is_none_or(|c| c.tags[slot] == 0)
                        {
                            continue;
                        }
                        let address = CellAddress {
                            sheet_id: range.sheet_id.clone(),
                            row,
                            column,
                        };
                        if let Some(cell) = get_cell(page, &address)? {
                            visitor(address, cell)?;
                        }
                    }
                }
            }
        }
        Ok(())
    }
}
impl WorkbookPages {
    fn sheet(&self, a: &CellAddress) -> Option<&SheetManifest> {
        self.manifest
            .sheets
            .iter()
            .find(|s| s.sheet_id == a.sheet_id && a.row < s.row_count && a.column < s.column_count)
    }
    fn sheet_id(&self, id: &str) -> KernelResult<&SheetManifest> {
        self.manifest
            .sheets
            .iter()
            .find(|s| s.sheet_id == id)
            .ok_or_else(|| KernelError::new("SHEET_NOT_FOUND", "Worksheet does not exist").at(id))
    }
}

fn validate_sheets(sheets: &[SheetManifest]) -> KernelResult<()> {
    let mut ids = BTreeSet::new();
    for s in sheets {
        if s.sheet_id.is_empty()
            || !ids.insert(s.sheet_id.clone())
            || s.row_count == 0
            || s.row_count > MAX_ROWS
            || s.column_count == 0
            || s.column_count > MAX_COLUMNS
        {
            return Err(KernelError::new(
                "SHEET_INVALID",
                "Invalid or duplicate worksheet manifest",
            ));
        }
    }
    Ok(())
}
fn validate_manifest(m: &WorkbookManifest) -> KernelResult<()> {
    if m.schema != "WorkbookManifest"
        || m.version != WORKBOOK_MANIFEST_VERSION
        || m.unit_id.is_empty()
    {
        return Err(KernelError::new(
            "MANIFEST_VERSION_UNSUPPORTED",
            "Only WorkbookManifest version 11 is accepted",
        ));
    }
    validate_sheets(&m.sheets)?;
    let mut keys = BTreeSet::new();
    for d in &m.pages {
        let key = d.key();
        if !keys.insert(key.clone())
            || d.byte_length == 0
            || d.byte_length > 1024 * 1024
            || d.cell_count == 0
            || d.checksum.len() != 64
            || !d.checksum.bytes().all(|b| b.is_ascii_hexdigit())
        {
            return Err(KernelError::new(
                "MANIFEST_INVALID",
                "Invalid page directory",
            ));
        }
        let s = m
            .sheets
            .iter()
            .find(|s| s.sheet_id == key.sheet_id)
            .ok_or_else(|| KernelError::new("MANIFEST_INVALID", "Page references unknown sheet"))?;
        let row_start = key
            .page_row
            .checked_mul(PAGE_ROWS)
            .ok_or_else(|| KernelError::new("MANIFEST_INVALID", "Page row coordinate overflows"))?;
        let col_start = key.page_column.checked_mul(PAGE_COLUMNS).ok_or_else(|| {
            KernelError::new("MANIFEST_INVALID", "Page column coordinate overflows")
        })?;
        if row_start >= s.row_count
            || col_start >= s.column_count
            || d.cell_count > PAGE_ROWS * PAGE_COLUMNS
        {
            return Err(KernelError::new(
                "MANIFEST_INVALID",
                "Page is outside worksheet bounds",
            ));
        }
        if let Some(range) = &d.occupied_range {
            range.validate()?;
            if range.sheet_id != key.sheet_id
                || range.start_row < row_start
                || range.end_row >= (row_start + PAGE_ROWS).min(s.row_count)
                || range.start_column < col_start
                || range.end_column >= (col_start + PAGE_COLUMNS).min(s.column_count)
            {
                return Err(KernelError::new(
                    "MANIFEST_INVALID",
                    "Page occupied range is outside page bounds",
                ));
            }
        } else {
            return Err(KernelError::new(
                "MANIFEST_INVALID",
                "Non-empty page must have an occupied range",
            ));
        }
    }
    Ok(())
}
fn validate_writes(sheets: &[SheetManifest], writes: &[CellWrite]) -> KernelResult<()> {
    for w in writes {
        w.address.validate()?;
        let _sheet = sheets
            .iter()
            .find(|s| {
                s.sheet_id == w.address.sheet_id
                    && w.address.row < s.row_count
                    && w.address.column < s.column_count
            })
            .ok_or_else(|| {
                KernelError::new(
                    "CELL_ADDRESS_INVALID",
                    "Cell address is outside the worksheet",
                )
                .at(format!(
                    "{}:{}:{}",
                    w.address.sheet_id, w.address.row, w.address.column
                ))
            })?;
        if let Some(c) = &w.cell {
            c.value.validate()?;
            if let Some(f) = &c.formula {
                if f.len() > 1_048_576 {
                    return Err(KernelError::new(
                        "CELL_VALUE_INVALID",
                        "Formula is too large",
                    ));
                }
            }
        }
    }
    Ok(())
}

fn set_cell(page: &mut CellPage, address: &CellAddress, cell: Option<&Cell>) -> KernelResult<()> {
    let col = (address.column % PAGE_COLUMNS) as usize;
    let row = (address.row % PAGE_ROWS) as usize;
    let Some(cell) = cell else {
        if let Some(column) = &mut page.columns[col] {
            column.tags[row] = 0;
            column.numbers[row] = 0.0;
            column.text_ids[row] = 0;
            column.formula_ids[row] = 0;
            column.metadata_ids[row] = 0;
        }
        return Ok(());
    };
    let value_ref = &cell.value;
    let (tag, value_id, number) = match value_ref {
        Scalar::Null => (1, 0, 0.0),
        Scalar::Boolean(v) => (if *v { 3 } else { 2 }, 0, 0.0),
        Scalar::Number(v) => (4, 0, *v),
        Scalar::Text(v) => (5, dict_string(page, v), 0.0),
        Scalar::Error(v) => (6, dict_error(page, v), 0.0),
    };
    let formula_id = cell
        .formula
        .as_deref()
        .map(|f| dict_formula(page, f) + 1)
        .unwrap_or(0);
    let metadata_id = if cell.metadata.is_empty() {
        0
    } else {
        page.metadata_dictionary.push(cell.metadata.clone());
        page.metadata_dictionary.len() as u32
    };
    let column = page.columns[col].get_or_insert_with(Column::default);
    column.tags[row] = tag;
    column.numbers[row] = number;
    column.text_ids[row] = value_id;
    column.formula_ids[row] = formula_id;
    column.metadata_ids[row] = metadata_id;
    Ok(())
}
fn get_cell(page: &CellPage, address: &CellAddress) -> KernelResult<Option<Cell>> {
    let Some(c) = page.columns[(address.column % PAGE_COLUMNS) as usize].as_ref() else {
        return Ok(None);
    };
    let r = (address.row % PAGE_ROWS) as usize;
    let value = match c.tags[r] {
        0 => return Ok(None),
        1 => Scalar::Null,
        2 => Scalar::Boolean(false),
        3 => Scalar::Boolean(true),
        4 => Scalar::Number(c.numbers[r]),
        5 => Scalar::Text(
            page.text_dictionary
                .get(c.text_ids[r] as usize)
                .ok_or_else(|| {
                    KernelError::new("PAGE_ENCODING_INVALID", "Text dictionary index is invalid")
                })?
                .clone(),
        ),
        6 => Scalar::Error(
            page.errors
                .get(c.text_ids[r] as usize)
                .ok_or_else(|| {
                    KernelError::new("PAGE_ENCODING_INVALID", "Error dictionary index is invalid")
                })?
                .clone(),
        ),
        _ => {
            return Err(KernelError::new(
                "PAGE_ENCODING_INVALID",
                "Unknown cell tag",
            ));
        }
    };
    let formula = if c.formula_ids[r] == 0 {
        None
    } else {
        Some(
            page.formula_dictionary
                .get((c.formula_ids[r] - 1) as usize)
                .ok_or_else(|| {
                    KernelError::new(
                        "PAGE_ENCODING_INVALID",
                        "Formula dictionary index is invalid",
                    )
                })?
                .clone(),
        )
    };
    let metadata = if c.metadata_ids[r] == 0 {
        BTreeMap::new()
    } else {
        page.metadata_dictionary
            .get((c.metadata_ids[r] - 1) as usize)
            .ok_or_else(|| {
                KernelError::new(
                    "PAGE_ENCODING_INVALID",
                    "Metadata dictionary index is invalid",
                )
            })?
            .clone()
    };
    Ok(Some(Cell {
        value,
        formula,
        metadata,
    }))
}
fn page_is_empty(page: &CellPage) -> bool {
    page.columns
        .iter()
        .all(|c| c.as_ref().map_or(true, |c| c.tags.iter().all(|t| *t == 0)))
}
fn page_stats(page: &CellPage, key: &PageKey) -> (u32, Option<RangeRef>) {
    let mut count = 0u32;
    let mut range: Option<RangeRef> = None;
    for (column_index, column) in page.columns.iter().enumerate() {
        let Some(column) = column else { continue };
        for (row_index, tag) in column.tags.iter().enumerate() {
            if *tag == 0 {
                continue;
            }
            count += 1;
            let row = key.page_row * PAGE_ROWS + row_index as u32;
            let column = key.page_column * PAGE_COLUMNS + column_index as u32;
            range = Some(match range {
                None => RangeRef {
                    sheet_id: key.sheet_id.clone(),
                    start_row: row,
                    end_row: row,
                    start_column: column,
                    end_column: column,
                },
                Some(mut current) => {
                    current.start_row = current.start_row.min(row);
                    current.end_row = current.end_row.max(row);
                    current.start_column = current.start_column.min(column);
                    current.end_column = current.end_column.max(column);
                    current
                }
            });
        }
    }
    (count, range)
}
fn dict_string(page: &mut CellPage, value: &str) -> u32 {
    if let Some(i) = page.text_index.get(value) {
        *i
    } else {
        let i = page.text_dictionary.len() as u32;
        page.text_dictionary.push(value.into());
        page.text_index.insert(value.into(), i);
        i
    }
}
fn dict_formula(page: &mut CellPage, value: &str) -> u32 {
    if let Some(i) = page.formula_index.get(value) {
        *i
    } else {
        let i = page.formula_dictionary.len() as u32;
        page.formula_dictionary.push(value.into());
        page.formula_index.insert(value.into(), i);
        i
    }
}
fn dict_error(page: &mut CellPage, value: &FormulaError) -> u32 {
    let key = (
        value.kind.clone(),
        value.code.clone(),
        value.message.clone(),
    );
    if let Some(i) = page.error_index.get(&key) {
        *i
    } else {
        let i = page.errors.len() as u32;
        page.errors.push(value.clone());
        page.error_index.insert(key, i);
        i
    }
}

#[derive(Serialize, Deserialize)]
struct PageDictionaries {
    text: Vec<String>,
    formula: Vec<String>,
    metadata: Vec<BTreeMap<String, Value>>,
    errors: Vec<FormulaError>,
}
fn encode_page(page: &CellPage) -> KernelResult<Vec<u8>> {
    let dict = serde_json::to_vec(&PageDictionaries {
        text: page.text_dictionary.clone(),
        formula: page.formula_dictionary.clone(),
        metadata: page.metadata_dictionary.clone(),
        errors: page.errors.clone(),
    })
    .map_err(|e| KernelError::new("PAGE_ENCODING_INVALID", e.to_string()))?;
    let mut out = Vec::with_capacity(16 + 32 * 1024);
    out.extend_from_slice(PAGE_MAGIC);
    out.push(PAGE_ENCODING_VERSION);
    out.extend_from_slice(&(dict.len() as u32).to_le_bytes());
    out.extend_from_slice(&dict);
    for col in &page.columns {
        let col = col.as_ref().cloned().unwrap_or_default();
        out.extend_from_slice(&col.tags);
        for n in &col.numbers {
            out.extend_from_slice(&n.to_le_bytes());
        }
        for ids in [&col.text_ids, &col.formula_ids, &col.metadata_ids] {
            for id in ids.iter() {
                out.extend_from_slice(&id.to_le_bytes());
            }
        }
    }
    Ok(out)
}
fn decode_page(bytes: &[u8]) -> KernelResult<CellPage> {
    if bytes.len() < 9 || &bytes[0..4] != PAGE_MAGIC {
        return Err(KernelError::new(
            "PAGE_ENCODING_INVALID",
            "Page magic is invalid",
        ));
    }
    if bytes[4] != PAGE_ENCODING_VERSION {
        return Err(KernelError::new(
            "PAGE_ENCODING_UNSUPPORTED",
            "Page encoding version is unsupported",
        ));
    }
    let dl = u32::from_le_bytes(bytes[5..9].try_into().unwrap()) as usize;
    let mut p = 9;
    if dl > bytes.len() - p {
        return Err(KernelError::new(
            "PAGE_ENCODING_INVALID",
            "Dictionary length is invalid",
        ));
    }
    let dict: PageDictionaries = serde_json::from_slice(&bytes[p..p + dl])
        .map_err(|e| KernelError::new("PAGE_ENCODING_INVALID", e.to_string()))?;
    p += dl;
    let expected = 32 * (1024 + 8 * 1024 + 3 * 4 * 1024);
    if bytes.len() - p != expected {
        return Err(KernelError::new(
            "PAGE_ENCODING_INVALID",
            "Column data length is invalid",
        ));
    }
    let mut page = CellPage {
        columns: Vec::new(),
        text_dictionary: dict.text,
        formula_dictionary: dict.formula,
        metadata_dictionary: dict.metadata,
        errors: dict.errors,
        text_index: HashMap::new(),
        formula_index: HashMap::new(),
        error_index: HashMap::new(),
    };
    for (i, value) in page.text_dictionary.iter().enumerate() {
        if page.text_index.insert(value.clone(), i as u32).is_some() {
            return Err(KernelError::new(
                "PAGE_ENCODING_INVALID",
                "Text dictionary contains duplicate entries",
            ));
        }
    }
    for (i, value) in page.formula_dictionary.iter().enumerate() {
        if page.formula_index.insert(value.clone(), i as u32).is_some() {
            return Err(KernelError::new(
                "PAGE_ENCODING_INVALID",
                "Formula dictionary contains duplicate entries",
            ));
        }
    }
    for (i, value) in page.errors.iter().enumerate() {
        if !matches!(value.kind.as_str(), "error") {
            return Err(KernelError::new(
                "PAGE_ENCODING_INVALID",
                "Invalid formula error kind",
            ));
        }
        let key = (
            value.kind.clone(),
            value.code.clone(),
            value.message.clone(),
        );
        if page.error_index.insert(key, i as u32).is_some() {
            return Err(KernelError::new(
                "PAGE_ENCODING_INVALID",
                "Error dictionary contains duplicate entries",
            ));
        }
    }
    for _ in 0..32 {
        let mut c = Column::default();
        c.tags.copy_from_slice(&bytes[p..p + 1024]);
        p += 1024;
        for tag in &c.tags {
            if *tag > 6 {
                return Err(KernelError::new(
                    "PAGE_ENCODING_INVALID",
                    "Unknown cell tag",
                ));
            }
        }
        for n in &mut c.numbers {
            *n = f64::from_le_bytes(bytes[p..p + 8].try_into().unwrap());
            p += 8;
            if !n.is_finite() {
                return Err(KernelError::new(
                    "PAGE_ENCODING_INVALID",
                    "Non-finite number in page",
                ));
            }
        }
        for ids in [&mut c.text_ids, &mut c.formula_ids, &mut c.metadata_ids] {
            for id in ids.iter_mut() {
                *id = u32::from_le_bytes(bytes[p..p + 4].try_into().unwrap());
                p += 4;
            }
        }
        for row in 0..1024 {
            if c.tags[row] == 5 && c.text_ids[row] as usize >= page.text_dictionary.len() {
                return Err(KernelError::new(
                    "PAGE_ENCODING_INVALID",
                    "Text dictionary index is invalid",
                ));
            }
            if c.tags[row] == 6 && c.text_ids[row] as usize >= page.errors.len() {
                return Err(KernelError::new(
                    "PAGE_ENCODING_INVALID",
                    "Error dictionary index is invalid",
                ));
            }
            if c.formula_ids[row] > 0
                && (c.formula_ids[row] - 1) as usize >= page.formula_dictionary.len()
            {
                return Err(KernelError::new(
                    "PAGE_ENCODING_INVALID",
                    "Formula dictionary index is invalid",
                ));
            }
            if c.metadata_ids[row] > 0
                && (c.metadata_ids[row] - 1) as usize >= page.metadata_dictionary.len()
            {
                return Err(KernelError::new(
                    "PAGE_ENCODING_INVALID",
                    "Metadata dictionary index is invalid",
                ));
            }
        }
        page.columns.push(if c.tags.iter().all(|tag| *tag == 0) {
            None
        } else {
            Some(c)
        });
    }
    Ok(page)
}
fn checksum(bytes: &[u8]) -> String {
    let mut h = Sha256::new();
    h.update(bytes);
    let digest = h.finalize();
    digest.iter().map(|b| format!("{:02x}", b)).collect()
}
fn b64_encode(bytes: &[u8]) -> String {
    const T: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut s = String::new();
    for chunk in bytes.chunks(3) {
        let n = ((chunk[0] as u32) << 16)
            | ((chunk.get(1).copied().unwrap_or(0) as u32) << 8)
            | chunk.get(2).copied().unwrap_or(0) as u32;
        s.push(T[((n >> 18) & 63) as usize] as char);
        s.push(T[((n >> 12) & 63) as usize] as char);
        s.push(if chunk.len() > 1 {
            T[((n >> 6) & 63) as usize] as char
        } else {
            '='
        });
        s.push(if chunk.len() > 2 {
            T[(n & 63) as usize] as char
        } else {
            '='
        });
    }
    s
}
fn b64_decode(s: &str) -> KernelResult<Vec<u8>> {
    fn v(c: u8) -> Option<u8> {
        match c {
            b'A'..=b'Z' => Some(c - b'A'),
            b'a'..=b'z' => Some(c - b'a' + 26),
            b'0'..=b'9' => Some(c - b'0' + 52),
            b'+' => Some(62),
            b'/' => Some(63),
            _ => None,
        }
    }
    let b = s.as_bytes();
    if b.len() % 4 != 0 {
        return Err(KernelError::new(
            "PAGE_PAYLOAD_INVALID",
            "Base64 length is invalid",
        ));
    }
    let mut out = Vec::new();
    for c in b.chunks(4) {
        let a = v(c[0]).ok_or_else(|| KernelError::new("PAGE_PAYLOAD_INVALID", "Invalid Base64"))?
            as u32;
        let d = v(c[1]).ok_or_else(|| KernelError::new("PAGE_PAYLOAD_INVALID", "Invalid Base64"))?
            as u32;
        let e = if c[2] == b'=' {
            0
        } else {
            v(c[2]).ok_or_else(|| KernelError::new("PAGE_PAYLOAD_INVALID", "Invalid Base64"))?
                as u32
        };
        let f = if c[3] == b'=' {
            0
        } else {
            v(c[3]).ok_or_else(|| KernelError::new("PAGE_PAYLOAD_INVALID", "Invalid Base64"))?
                as u32
        };
        let n = (a << 18) | (d << 12) | (e << 6) | f;
        out.push((n >> 16) as u8);
        if c[2] != b'=' {
            out.push((n >> 8) as u8);
        }
        if c[3] != b'=' {
            out.push(n as u8);
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn store() -> WorkbookPages {
        WorkbookPages::create(
            "u",
            "book",
            vec![SheetManifest {
                sheet_id: "s".into(),
                name: "Sheet".into(),
                row_count: 1_048_576,
                column_count: 16_384,
                metadata: BTreeMap::new(),
            }],
        )
        .unwrap()
    }
    fn address(row: u32, column: u32) -> CellAddress {
        CellAddress {
            sheet_id: "s".into(),
            row,
            column,
        }
    }

    #[test]
    fn new_workbooks_publish_complete_canonical_metadata() {
        let manifest = store().manifest();
        for key in [
            "dateSystem",
            "numericContext",
            "collationContext",
            "calculationSettings",
            "dimensionMetrics",
            "editingOptions",
            "definedNameModels",
            "dataModel",
        ] {
            assert!(
                manifest.metadata.contains_key(key),
                "missing canonical metadata: {key}"
            );
        }
    }

    #[test]
    fn current_region_resolves_in_kernel_and_rejects_unloaded_pages() {
        let mut pages = store();
        let committed = pages
            .apply_writes(
                "region",
                0,
                vec![
                    CellWrite {
                        address: address(0, 0),
                        cell: Some(Cell {
                            value: Scalar::Text("a".into()),
                            ..Cell::default()
                        }),
                    },
                    CellWrite {
                        address: address(0, 1),
                        cell: Some(Cell {
                            value: Scalar::Text("b".into()),
                            ..Cell::default()
                        }),
                    },
                    CellWrite {
                        address: address(1, 0),
                        cell: Some(Cell {
                            value: Scalar::Text("c".into()),
                            ..Cell::default()
                        }),
                    },
                    CellWrite {
                        address: address(3, 3),
                        cell: Some(Cell {
                            value: Scalar::Text("outside".into()),
                            ..Cell::default()
                        }),
                    },
                ],
            )
            .unwrap();
        assert_eq!(
            pages.resolve_current_region("s", 0, 0).unwrap(),
            RangeRef {
                sheet_id: "s".into(),
                start_row: 0,
                end_row: 1,
                start_column: 0,
                end_column: 1,
            }
        );
        let unloaded = WorkbookPages::open(committed.manifest).unwrap();
        assert_eq!(
            unloaded.resolve_current_region("s", 0, 0).unwrap_err().code,
            "DATA_PAGE_UNAVAILABLE"
        );
    }

    #[test]
    fn page_round_trip_and_boundary() {
        let mut pages = store();
        pages
            .apply_writes(
                "op",
                0,
                vec![CellWrite {
                    address: address(1024, 32),
                    cell: Some(Cell {
                        value: Scalar::Text("x".into()),
                        ..Cell::default()
                    }),
                }],
            )
            .unwrap();
        assert_eq!(
            pages.read_cell(&address(1024, 32)).unwrap().unwrap().value,
            Scalar::Text("x".into())
        );
        assert_eq!(pages.read_cell(&address(1023, 31)).unwrap(), None);
    }

    #[test]
    fn stale_and_budget_reject_without_mutation() {
        let mut pages = store();
        pages.set_page_budget(1).unwrap();
        let before = pages.manifest();
        let result = pages.apply_writes(
            "op",
            0,
            vec![CellWrite {
                address: address(0, 0),
                cell: Some(Cell {
                    value: Scalar::Number(1.0),
                    ..Cell::default()
                }),
            }],
        );
        assert_eq!(result.unwrap_err().code, "PAGE_BUDGET_EXCEEDED");
        assert_eq!(pages.manifest(), before);
        assert_eq!(
            pages.apply_writes("stale", 1, Vec::new()).unwrap_err().code,
            "STALE_REVISION"
        );
    }

    #[test]
    fn unavailable_manifest_page_is_not_blank() {
        let mut pages = store();
        pages
            .apply_writes(
                "op",
                0,
                vec![CellWrite {
                    address: address(0, 0),
                    cell: Some(Cell {
                        value: Scalar::Boolean(true),
                        ..Cell::default()
                    }),
                }],
            )
            .unwrap();
        let manifest = pages.manifest();
        let mut reopened = WorkbookPages::open(manifest).unwrap();
        assert_eq!(
            reopened.read_cell(&address(0, 0)).unwrap_err().code,
            "DATA_PAGE_UNAVAILABLE"
        );
        let key = PageKey::for_address(&address(0, 0));
        let descriptor = reopened
            .manifest()
            .pages
            .into_iter()
            .find(|d| d.key() == key)
            .unwrap();
        let bytes = pages.page_bytes(&key).unwrap();
        reopened.load_page(&descriptor, &bytes).unwrap();
        assert_eq!(
            reopened.read_cell(&address(0, 0)).unwrap().unwrap().value,
            Scalar::Boolean(true)
        );
    }

    #[test]
    fn metadata_and_writes_share_one_revision() {
        let mut pages = store();
        let mut next = pages.manifest();
        next.name = "renamed".into();
        next.metadata
            .insert("source".into(), Value::String("test".into()));
        let change = pages
            .apply_transaction(
                "op",
                0,
                vec![CellWrite {
                    address: address(0, 0),
                    cell: Some(Cell {
                        value: Scalar::Number(7.0),
                        ..Cell::default()
                    }),
                }],
                next,
            )
            .unwrap();
        assert_eq!(change.revision, 1);
        assert_eq!(change.manifest.name, "renamed");
        assert_eq!(pages.revision(), 1);
        assert_eq!(
            pages.read_cell(&address(0, 0)).unwrap().unwrap().value,
            Scalar::Number(7.0)
        );
    }
}
