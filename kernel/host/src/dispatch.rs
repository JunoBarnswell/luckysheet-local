use base64::{Engine as _, engine::general_purpose::STANDARD};
use kernel_commands::{AccessRole, CommandRequest};
use kernel_core::*;
use kernel_formula::{DefinedNameScope, FormulaRuntime, FormulaTable, InspectionQuery};
use kernel_geometry::{GeometryRequest, HeaderAxis, Point};
use serde::{Deserialize, Serialize, de::DeserializeOwned};
use serde_json::{Value, json};
use std::collections::BTreeMap;
use std::sync::{Arc, atomic::AtomicBool};

pub const MAX_FRAME_BYTES: usize = 16 * 1024 * 1024;
const MAX_RANGE_RESPONSE_CELLS: usize = 65_536;
/// A host is a process-wide runtime, so every registry needs a hard bound.
/// Callers must close an unused workbook before opening another one.
const MAX_CONTEXTS: usize = 64;

#[cfg(test)]
#[path = "analytics_tests.rs"]
mod analytics_tests;
#[cfg(test)]
#[path = "formula_tests.rs"]
mod formula_tests;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct Invocation {
    protocol_version: u32,
    request_id: String,
    operation: String,
    params: Value,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct FormulaOverride {
    address: CellAddress,
    value: Scalar,
}

#[derive(Default)]
pub struct KernelHost {
    workbooks: BTreeMap<String, WorkbookPages>,
    formulas: BTreeMap<String, FormulaRuntime>,
    analytics: BTreeMap<String, kernel_analytics::AnalyticsRuntime>,
    cancellation: Arc<AtomicBool>,
}

impl KernelHost {
    fn ensure_context_capacity(
        current: usize,
        unit_id: &str,
        context: &str,
    ) -> KernelResult<()> {
        if current >= MAX_CONTEXTS {
            return Err(KernelError::new(
                "KERNEL_CONTEXT_LIMIT",
                format!("The kernel host has reached its {context} context limit"),
            )
            .at(unit_id)
            .recover("close-unused-workbook-context"));
        }
        Ok(())
    }

    fn ensure_formula_runtime(&mut self, unit_id: &str) -> KernelResult<()> {
        if !self.formulas.contains_key(unit_id) {
            let workbook = self.workbooks.get(unit_id).ok_or_else(|| {
                KernelError::new("WORKBOOK_NOT_OPEN", "Workbook is not open").at(unit_id)
            })?;
            Self::ensure_context_capacity(self.formulas.len(), unit_id, "formula")?;
            self.formulas
                .insert(unit_id.to_owned(), build_formula_runtime(workbook)?);
        }
        Ok(())
    }
    pub fn invoke_cancellable(&mut self, bytes: &[u8], cancellation: Arc<AtomicBool>) -> Vec<u8> {
        self.cancellation = cancellation;
        let response = self.invoke(bytes);
        self.cancellation = Arc::new(AtomicBool::new(false));
        response
    }
    pub fn invoke(&mut self, bytes: &[u8]) -> Vec<u8> {
        let mut request_id = String::new();
        let outcome = (|| {
            if bytes.len() > MAX_FRAME_BYTES {
                return Err(KernelError::new(
                    "KERNEL_PAYLOAD_TOO_LARGE",
                    "Control frame exceeds 16 MiB",
                ));
            }
            let request: Invocation = serde_json::from_slice(bytes)
                .map_err(|e| KernelError::new("KERNEL_PROTOCOL_ERROR", e.to_string()))?;
            request_id = request.request_id.clone();
            if request.protocol_version != KERNEL_PROTOCOL_VERSION || request_id.is_empty() {
                return Err(KernelError::new(
                    "KERNEL_VERSION_MISMATCH",
                    "Protocol version 1 and nonempty requestId are required",
                ));
            }
            self.dispatch(&request.operation, request.params)
        })();
        let envelope = match outcome {
            Ok(result) => {
                json!({"protocolVersion":KERNEL_PROTOCOL_VERSION,"requestId":request_id,"ok":true,"result":result})
            }
            Err(error) => {
                json!({"protocolVersion":KERNEL_PROTOCOL_VERSION,"requestId":request_id,"ok":false,"error":error})
            }
        };
        let bytes = serde_json::to_vec(&envelope)
            .expect("kernel envelopes contain only serializable JSON values");
        if bytes.len() <= MAX_FRAME_BYTES {
            bytes
        } else {
            serde_json::to_vec(&json!({"protocolVersion":1,"requestId":request_id,"ok":false,"error":KernelError::new("KERNEL_PAYLOAD_TOO_LARGE","Result must be read in bounded pages").recover("use-data-pages")})).unwrap()
        }
    }

    pub(crate) fn workbook(&self, params: &Value) -> KernelResult<&WorkbookPages> {
        let id = string(params, "unitId")?;
        let workbook = self
            .workbooks
            .get(id)
            .ok_or_else(|| KernelError::new("WORKBOOK_NOT_OPEN", "Workbook is not open").at(id))?;
        let revision = number(params, "revision")?;
        if workbook.revision() != revision {
            return Err(KernelError::new(
                "STALE_REVISION",
                "Requested revision does not match the open workbook",
            )
            .at(id)
            .recover("refresh-manifest"));
        }
        Ok(workbook)
    }

    pub fn dispatch(&mut self, operation: &str, params: Value) -> KernelResult<Value> {
        match operation {
            "init" => Ok(
                json!({"protocolVersion":1,"manifestVersion":11,"operations":["init","open","create","manifest","sheet.stats","cell.get","range.get","dataRegion.resolve","page.get","page.load","command","close","formula.functions","formula.evaluate","formula.recalculate","formula.inspect","formula.trace","formula.spillValue","analytics.execute","geometry.computePaneMap","geometry.hitTest","geometry.cellRect","geometry.headerRect"]}),
            ),
            "create" => {
                let unit_id = string(&params, "unitId")?.to_owned();
                if self.workbooks.contains_key(&unit_id) {
                    return Err(KernelError::new(
                        "WORKBOOK_ALREADY_OPEN",
                        "Workbook identity is already open",
                    )
                    .at(unit_id));
                }
                let sheets: Vec<SheetManifest> = match params.get("sheets") {
                    Some(value) => decode(value.clone())?,
                    None => vec![SheetManifest {
                        sheet_id: "sheet-1".into(),
                        name: "Sheet1".into(),
                        row_count: MAX_ROWS,
                        column_count: MAX_COLUMNS,
                        metadata: BTreeMap::new(),
                    }],
                };
                if sheets.is_empty() {
                    return Err(KernelError::new(
                        "SHEET_INVALID",
                        "A new workbook requires at least one worksheet",
                    ));
                }
                let workbook = WorkbookPages::create(&unit_id, string(&params, "name")?, sheets)?;
                Self::ensure_context_capacity(self.workbooks.len(), &unit_id, "workbook")?;
                let manifest = encode(workbook.manifest())?;
                self.workbooks.insert(unit_id, workbook);
                Ok(manifest)
            }
            "open" => {
                let manifest: WorkbookManifest = decode(required(&params, "manifest")?.clone())?;
                if !self.workbooks.contains_key(&manifest.unit_id) {
                    Self::ensure_context_capacity(
                        self.workbooks.len(),
                        &manifest.unit_id,
                        "workbook",
                    )?;
                }
                let mut workbook = match self.workbooks.get(&manifest.unit_id) {
                    Some(previous) => WorkbookPages::open_reusing(manifest.clone(), previous)?,
                    None => WorkbookPages::open(manifest.clone())?,
                };
                if let Some(pages) = params.get("pages") {
                    for page in decode::<Vec<PagePayload>>(pages.clone())? {
                        workbook.load_page(
                            &page.descriptor,
                            &decode_page_payload(&page.payload_base64)?,
                        )?;
                    }
                }
                self.formulas.remove(&manifest.unit_id);
                self.analytics.remove(&manifest.unit_id);
                self.workbooks.insert(manifest.unit_id.clone(), workbook);
                Ok(
                    json!({"unitId":manifest.unit_id,"revision":manifest.revision,"pageCount":manifest.pages.len()}),
                )
            }
            "restore" => {
                let id = string(&params, "unitId")?;
                let role: AccessRole = decode(required(&params, "accessRole")?.clone())?;
                if role != AccessRole::Owner {
                    return Err(KernelError::new(
                        "FORBIDDEN",
                        "Restore requires workbook owner access",
                    ));
                }
                let workbook = self.workbooks.get(id).ok_or_else(|| {
                    KernelError::new("WORKBOOK_NOT_OPEN", "Restore workbook is not open").at(id)
                })?;
                let mut staged = workbook.clone();
                let mut changes = staged.restore_manifest(
                    string(&params, "operationId")?.to_owned(),
                    number(&params, "baseRevision")?,
                    decode(required(&params, "targetManifest")?.clone())?,
                )?;
                changes.history.required_role = AccessRole::Owner;
                let result = encode(changes)?;
                if serde_json::to_vec(&result)
                    .map_err(|e| KernelError::new("KERNEL_RESPONSE_INVALID", e.to_string()))?
                    .len()
                    + 1024
                    > MAX_FRAME_BYTES
                {
                    return Err(KernelError::new(
                        "KERNEL_PAYLOAD_TOO_LARGE",
                        "Restore result exceeds the control frame budget",
                    ));
                }
                self.workbooks.insert(id.to_owned(), staged);
                self.formulas.remove(id);
                self.analytics.remove(id);
                Ok(result)
            }
            "copy" => {
                let source_id = string(&params, "sourceUnitId")?;
                let source_revision = number(&params, "sourceRevision")?;
                let target_id = string(&params, "targetUnitId")?;
                if self.workbooks.contains_key(target_id) {
                    return Err(KernelError::new(
                        "WORKBOOK_ALREADY_OPEN",
                        "Copy target is already open",
                    )
                    .at(target_id));
                }
                let source = self.workbooks.get(source_id).ok_or_else(|| {
                    KernelError::new("WORKBOOK_NOT_OPEN", "Copy source is not open").at(source_id)
                })?;
                if source.revision() != source_revision {
                    return Err(KernelError::new(
                        "STALE_REVISION",
                        "Copy source revision is stale",
                    )
                    .at(source_id));
                }
                Self::ensure_context_capacity(self.workbooks.len(), target_id, "workbook")?;
                let mut manifest = source.manifest();
                manifest.unit_id = target_id.to_owned();
                manifest.name = string(&params, "name")?.to_owned();
                manifest.revision = 0;
                for page in &mut manifest.pages {
                    page.revision = 0;
                }
                let copied = WorkbookPages::open(manifest.clone())?;
                let result = encode(json!({"manifest":manifest}))?;
                self.workbooks.insert(target_id.to_owned(), copied);
                Ok(result)
            }
            "manifest" => encode(self.workbook(&params)?.manifest()),
            "sheet.stats" => encode(
                self.workbook(&params)?
                    .sheet_stats(string(&params, "sheetId")?)?,
            ),
            "cell.get" => {
                let workbook = self.workbook(&params)?;
                let address: CellAddress = decode(required(&params, "address")?.clone())?;
                Ok(json!({"revision":workbook.revision(),"cell":workbook.read_cell(&address)?}))
            }
            "range.get" => {
                let workbook = self.workbook(&params)?;
                let range: RangeRef = decode(required(&params, "range")?.clone())?;
                let mut cells = Vec::new();
                workbook.read_range(&range, &mut |address, cell| {
                    if cells.len() >= MAX_RANGE_RESPONSE_CELLS {
                        return Err(KernelError::new(
                            "KERNEL_RANGE_TOO_LARGE",
                            "Read data pages for large ranges",
                        )
                        .recover("use-data-pages"));
                    }
                    cells.push(json!({"address":address,"cell":cell}));
                    Ok(())
                })?;
                Ok(json!({"revision":workbook.revision(),"cells":cells}))
            }
            "dataRegion.resolve" => {
                let workbook = self.workbook(&params)?;
                let active_row = u32::try_from(number(&params, "activeRow")?).map_err(|_| {
                    KernelError::new(
                        "CELL_ADDRESS_INVALID",
                        "Current-region row exceeds the kernel coordinate range",
                    )
                })?;
                let active_column =
                    u32::try_from(number(&params, "activeColumn")?).map_err(|_| {
                        KernelError::new(
                            "CELL_ADDRESS_INVALID",
                            "Current-region column exceeds the kernel coordinate range",
                        )
                    })?;
                let range = workbook.resolve_current_region(
                    string(&params, "sheetId")?,
                    active_row,
                    active_column,
                )?;
                Ok(json!({"revision":workbook.revision(),"range":range}))
            }
            "page.get" => {
                let workbook = self.workbook(&params)?;
                let key = page_key(&params)?;
                let manifest = workbook.manifest();
                let descriptor = manifest
                    .pages
                    .iter()
                    .find(|p| p.key() == key)
                    .ok_or_else(|| KernelError::new("PAGE_NOT_FOUND", "Page is not in manifest"))?
                    .clone();
                encode(PagePayload {
                    descriptor,
                    payload_base64: STANDARD.encode(workbook.page_bytes(&key)?),
                })
            }
            "page.load" => {
                self.workbook(&params)?;
                let page: PagePayload = decode(required(&params, "page")?.clone())?;
                let id = string(&params, "unitId")?;
                self.workbooks.get_mut(id).unwrap().load_page(
                    &page.descriptor,
                    &decode_page_payload(&page.payload_base64)?,
                )?;
                Ok(json!({"revision":number(&params,"revision")?,"loaded":page.descriptor}))
            }
            "command" => self.command(params),
            "command.policy" => {
                let request: CommandRequest = decode(params)?;
                encode(kernel_commands::required_role(&request)?)
            }
            "command.prepare" => {
                let id = string(&params, "unitId")?;
                let workbook = self.workbooks.get(id).ok_or_else(|| {
                    KernelError::new("WORKBOOK_NOT_OPEN", "Workbook is not open").at(id)
                })?;
                let role: AccessRole = decode(required(&params, "accessRole")?.clone())?;
                let mut intent = params.clone();
                intent.as_object_mut().unwrap().remove("accessRole");
                let request: CommandRequest = decode(intent)?;
                if role < kernel_commands::required_role(&request)? {
                    return Err(KernelError::new(
                        "FORBIDDEN",
                        "Workbook role does not permit this command",
                    ));
                }
                if request.base_revision != workbook.revision() {
                    return Err(KernelError::new(
                        "STALE_REVISION",
                        "Command base revision is stale",
                    ));
                }
                Ok(json!({"pages": kernel_commands::required_pages(workbook, &request)?}))
            }
            "close" => {
                let id = string(&params, "unitId")?;
                self.formulas.remove(id);
                self.analytics.remove(id);
                Ok(json!({"closed":self.workbooks.remove(id).is_some()}))
            }
            "formula.functions" => {
                Ok(json!({"functions": kernel_formula::function_capabilities()}))
            }
            "formula.evaluate" => {
                let id = string(&params, "unitId")?.to_owned();
                self.workbook(&params)?;
                self.ensure_formula_runtime(&id)?;
                let workbook = self.workbooks.get(&id).unwrap();
                let address: CellAddress = decode(required(&params, "address")?.clone())?;
                let overrides = params
                    .get("overrides")
                    .map(|value| decode::<Vec<FormulaOverride>>(value.clone()))
                    .transpose()?
                    .unwrap_or_default()
                    .into_iter()
                    .map(|item| (item.address, item.value))
                    .collect::<BTreeMap<_, _>>();
                let runtime = self.formulas.get(&id).unwrap();
                let result = if overrides.is_empty() {
                    runtime.evaluate(string(&params, "formula")?, &address, workbook)?
                } else {
                    runtime.evaluate_with_overrides(
                        string(&params, "formula")?,
                        &address,
                        workbook,
                        &overrides,
                    )?
                };
                Ok(json!({"revision":workbook.revision(),"value":result}))
            }
            "formula.recalculate" => self.recalculate(params),
            "formula.inspect" => {
                let id = string(&params, "unitId")?.to_owned();
                self.workbook(&params)?;
                self.ensure_formula_runtime(&id)?;
                let query: InspectionQuery = decode(params)?;
                encode(
                    self.formulas
                        .get(&id)
                        .unwrap()
                        .inspect_query(self.workbooks.get(&id).unwrap(), &query)?,
                )
            }
            "formula.trace" => {
                let id = string(&params, "unitId")?.to_owned();
                self.workbook(&params)?;
                self.ensure_formula_runtime(&id)?;
                let address: CellAddress = decode(required(&params, "address")?.clone())?;
                let trace = self
                    .formulas
                    .get(&id)
                    .unwrap()
                    .trace(&address, self.workbooks.get(&id).unwrap())?;
                Ok(
                    json!({"revision":self.workbooks.get(&id).unwrap().revision(),"value":trace.value,"steps":trace.steps}),
                )
            }
            "formula.spillValue" => {
                let id = string(&params, "unitId")?.to_owned();
                self.workbook(&params)?;
                self.ensure_formula_runtime(&id)?;
                let address: CellAddress = decode(required(&params, "address")?.clone())?;
                let spill = self
                    .formulas
                    .get(&id)
                    .unwrap()
                    .spill_value(&address, self.workbooks.get(&id).unwrap())?;
                Ok(
                    json!({"revision":self.workbooks.get(&id).unwrap().revision(),"value":spill.value,"isSpill":spill.is_spill}),
                )
            }
            "analytics.execute" => {
                self.workbook(&params)?;
                let request = required(&params, "request")?.clone();
                let id = string(&params, "unitId")?;
                if !self.analytics.contains_key(id) {
                    Self::ensure_context_capacity(self.analytics.len(), id, "analytics")?;
                }
                self.analytics.entry(id.to_owned()).or_default().execute(
                    request,
                    self.workbooks.get(id).unwrap(),
                    &self.cancellation,
                )
            }
            #[cfg(not(target_arch = "wasm32"))]
            "document.import" | "document.export" => {
                crate::native_document::dispatch(self, operation, params)
            }
            "geometry.computePaneMap" => encode(kernel_geometry::compute_pane_map(&decode::<
                GeometryRequest,
            >(
                params
            )?)?),
            "geometry.hitTest" => encode(kernel_geometry::hit_test(
                &decode(required(&params, "request")?.clone())?,
                decode::<Point>(required(&params, "point")?.clone())?,
            )?),
            "geometry.cellRect" => encode(kernel_geometry::cell_rect(
                &decode(required(&params, "request")?.clone())?,
                &decode(required(&params, "address")?.clone())?,
            )?),
            "geometry.headerRect" => encode(kernel_geometry::header_rect(
                &decode(required(&params, "request")?.clone())?,
                decode::<HeaderAxis>(required(&params, "axis")?.clone())?,
                params
                    .get("index")
                    .filter(|v| !v.is_null())
                    .map(|v| decode::<u32>(v.clone()))
                    .transpose()?,
            )?),
            _ => Err(KernelError::new(
                "UNSUPPORTED_FEATURE",
                format!("Unknown kernel operation: {operation}"),
            )
            .at(operation)),
        }
    }

    fn command(&mut self, params: Value) -> KernelResult<Value> {
        let id = string(&params, "unitId")?.to_owned();
        let access_role: AccessRole = decode(required(&params, "accessRole")?.clone())?;
        let mut intent = params.clone();
        intent
            .as_object_mut()
            .ok_or_else(|| {
                KernelError::new("KERNEL_REQUEST_INVALID", "Command params must be an object")
            })?
            .remove("accessRole");
        let request: CommandRequest = decode(intent)?;
        let workbook = self
            .workbooks
            .get(&id)
            .ok_or_else(|| KernelError::new("WORKBOOK_NOT_OPEN", "Workbook is not open").at(&id))?;
        let mut staged = workbook.clone();
        let changes = kernel_commands::execute_authorized(&mut staged, request, access_role)?;
        let result = encode(&changes)?;
        if serde_json::to_vec(&result)
            .map_err(|e| KernelError::new("KERNEL_RESPONSE_INVALID", e.to_string()))?
            .len()
            + 1024
            > MAX_FRAME_BYTES
        {
            return Err(KernelError::new(
                "KERNEL_PAYLOAD_TOO_LARGE",
                "ChangeSet must be staged in data pages before committing",
            )
            .recover("use-bulk-transaction"));
        }
        let staged_formula = if let Some(current) = self.formulas.get(&id) {
            if changes.history.metadata_after.is_some() {
                Some(build_formula_runtime(&staged)?)
            } else {
                let mut next = current.clone();
                for range in &changes.affected_ranges {
                    next.synchronize_range(range, &staged)?;
                }
                Some(next)
            }
        } else {
            None
        };
        self.workbooks.insert(id.clone(), staged);
        if let Some(runtime) = staged_formula {
            self.formulas.insert(id, runtime);
        }
        Ok(result)
    }

    fn recalculate(&mut self, params: Value) -> KernelResult<Value> {
        let id = string(&params, "unitId")?.to_owned();
        self.workbook(&params)?;
        let workbook = self.workbooks.get(&id).unwrap();
        let manifest = workbook.manifest();
        if !self.formulas.contains_key(&id) {
            self.formulas
                .insert(id.clone(), build_formula_runtime(workbook)?);
        }
        let values = if let Some(value) = params.get("address").filter(|value| !value.is_null()) {
            let address: CellAddress = decode(value.clone())?;
            self.formulas
                .get_mut(&id)
                .unwrap()
                .recalculate_cell(&address, workbook)?
                .into_iter()
                .map(|value| (address.clone(), value))
                .collect()
        } else {
            self.formulas.get_mut(&id).unwrap().recalculate(workbook)?
        };
        let generation = self.formulas.get(&id).unwrap().generation();
        let pending = self.formulas.get(&id).unwrap().dirty_count() != 0;
        Ok(json!({
            "revision":manifest.revision,
            "generation":generation,
            "recalculatedCount":values.len(),
            "pendingRecalculation":pending,
            "values":values.into_iter().map(|(address,value)|json!({"address":address,"value":value})).collect::<Vec<_>>()
        }))
    }
}

fn build_formula_runtime(workbook: &WorkbookPages) -> KernelResult<FormulaRuntime> {
    let manifest = workbook.manifest();
    let default_sheet = manifest
        .sheets
        .first()
        .ok_or_else(|| KernelError::new("WORKBOOK_INVALID", "No worksheets"))?
        .sheet_id
        .clone();
    let mut runtime = FormulaRuntime::new(default_sheet.clone());
    runtime.context.date1904 = manifest
        .metadata
        .get("dateSystem")
        .and_then(Value::as_str)
        .is_some_and(|value| value == "1904")
        || manifest
            .metadata
            .get("date1904")
            .and_then(Value::as_bool)
            .unwrap_or(false);
    for sheet in &manifest.sheets {
        runtime.register_sheet(&sheet.name, &sheet.sheet_id)?;
    }
    if let Some(names) = manifest.metadata.get("definedNameModels") {
        for value in names.as_array().ok_or_else(|| {
            KernelError::new("DEFINED_NAME_INVALID", "Defined names must be an array")
        })? {
            let name = required_text(value, "name", "DEFINED_NAME_INVALID")?;
            let formula = required_text(value, "formula", "DEFINED_NAME_INVALID")?;
            let scope = match value
                .get("scope")
                .and_then(Value::as_str)
                .unwrap_or("workbook")
            {
                "workbook" => DefinedNameScope::Workbook,
                "sheet" => DefinedNameScope::Sheet(
                    required_text(value, "sheetId", "DEFINED_NAME_INVALID")?.to_owned(),
                ),
                _ => {
                    return Err(KernelError::new(
                        "DEFINED_NAME_INVALID",
                        "Defined-name scope must be workbook or sheet",
                    )
                    .at(name));
                }
            };
            let scope_sheet = match &scope {
                DefinedNameScope::Workbook => default_sheet.clone(),
                DefinedNameScope::Sheet(sheet_id) => sheet_id.clone(),
            };
            runtime.define_name(
                name,
                formula,
                scope,
                &CellAddress {
                    sheet_id: scope_sheet,
                    row: 0,
                    column: 0,
                },
            )?;
        }
    }
    for sheet in &manifest.sheets {
        if let Some(tables) = sheet.metadata.get("sheetTables") {
            for value in tables.as_array().ok_or_else(|| {
                KernelError::new("TABLE_DEFINITION_INVALID", "Sheet tables must be an array")
            })? {
                let range: RangeRef = decode(
                    value
                        .get("range")
                        .ok_or_else(|| {
                            KernelError::new(
                                "TABLE_DEFINITION_INVALID",
                                "Sheet table range is required",
                            )
                        })?
                        .clone(),
                )?;
                if range.sheet_id != sheet.sheet_id {
                    return Err(KernelError::new(
                        "TABLE_DEFINITION_INVALID",
                        "Sheet table range belongs to another worksheet",
                    )
                    .at(required_text(
                        value,
                        "name",
                        "TABLE_DEFINITION_INVALID",
                    )?));
                }
                let columns = value
                    .get("columns")
                    .and_then(Value::as_array)
                    .ok_or_else(|| {
                        KernelError::new(
                            "TABLE_DEFINITION_INVALID",
                            "Sheet table columns must be an array",
                        )
                    })?
                    .iter()
                    .map(|column| {
                        required_text(column, "name", "TABLE_DEFINITION_INVALID").map(str::to_owned)
                    })
                    .collect::<KernelResult<Vec<_>>>()?;
                runtime.define_table(FormulaTable {
                    name: required_text(value, "name", "TABLE_DEFINITION_INVALID")?.to_owned(),
                    range,
                    has_header_row: value
                        .get("hasHeaderRow")
                        .and_then(Value::as_bool)
                        .ok_or_else(|| {
                            KernelError::new(
                                "TABLE_DEFINITION_INVALID",
                                "Sheet table hasHeaderRow is required",
                            )
                        })?,
                    has_total_row: value
                        .get("hasTotalRow")
                        .and_then(Value::as_bool)
                        .ok_or_else(|| {
                            KernelError::new(
                                "TABLE_DEFINITION_INVALID",
                                "Sheet table hasTotalRow is required",
                            )
                        })?,
                    columns,
                })?;
            }
        }
        let range = RangeRef {
            sheet_id: sheet.sheet_id.clone(),
            start_row: 0,
            end_row: sheet.row_count - 1,
            start_column: 0,
            end_column: sheet.column_count - 1,
        };
        workbook.read_range(&range, &mut |address, cell| {
            if let Some(formula) = cell.formula {
                runtime.set_formula(address, &formula)?;
            }
            Ok(())
        })?;
    }
    Ok(runtime)
}

fn required_text<'a>(value: &'a Value, key: &str, code: &str) -> KernelResult<&'a str> {
    value
        .get(key)
        .and_then(Value::as_str)
        .filter(|text| !text.trim().is_empty())
        .ok_or_else(|| KernelError::new(code, format!("{key} must be a nonempty string")))
}

fn required<'a>(params: &'a Value, key: &str) -> KernelResult<&'a Value> {
    params
        .get(key)
        .ok_or_else(|| KernelError::new("KERNEL_REQUEST_INVALID", format!("Missing field: {key}")))
}
fn string<'a>(params: &'a Value, key: &str) -> KernelResult<&'a str> {
    required(params, key)?
        .as_str()
        .filter(|v| !v.is_empty())
        .ok_or_else(|| {
            KernelError::new(
                "KERNEL_REQUEST_INVALID",
                format!("{key} must be a nonempty string"),
            )
        })
}
fn number(params: &Value, key: &str) -> KernelResult<u64> {
    required(params, key)?.as_u64().ok_or_else(|| {
        KernelError::new(
            "KERNEL_REQUEST_INVALID",
            format!("{key} must be an unsigned integer"),
        )
    })
}
fn decode<T: DeserializeOwned>(value: Value) -> KernelResult<T> {
    serde_json::from_value(value)
        .map_err(|e| KernelError::new("KERNEL_REQUEST_INVALID", e.to_string()))
}
fn encode<T: Serialize>(value: T) -> KernelResult<Value> {
    serde_json::to_value(value)
        .map_err(|e| KernelError::new("KERNEL_RESPONSE_INVALID", e.to_string()))
}
fn page_key(params: &Value) -> KernelResult<PageKey> {
    Ok(PageKey {
        sheet_id: string(params, "sheetId")?.into(),
        page_row: u32::try_from(number(params, "pageRow")?)
            .map_err(|_| KernelError::new("PAGE_ADDRESS_INVALID", "pageRow overflow"))?,
        page_column: u32::try_from(number(params, "pageColumn")?)
            .map_err(|_| KernelError::new("PAGE_ADDRESS_INVALID", "pageColumn overflow"))?,
    })
}
fn decode_page_payload(text: &str) -> KernelResult<Vec<u8>> {
    if text.len() > 1_400_000 {
        return Err(KernelError::new(
            "PAGE_BUDGET_EXCEEDED",
            "Page payload exceeds one MiB",
        ));
    }
    STANDARD
        .decode(text)
        .map_err(|e| KernelError::new("PAGE_ENCODING_INVALID", e.to_string()))
}
