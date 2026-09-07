//! The shared command owner. Commands stage reads, cells and metadata against
//! one revision; no reducer is permitted to mutate the committed page store.
use kernel_core::*;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::BTreeMap;

mod catalog;
mod cells;
mod objects;
mod preparation;
mod protection;
mod review;
mod sheet_metadata;
mod structure;
mod validation;
pub use catalog::MUTATION_IDS;
pub use preparation::required_pages;
#[cfg(test)]
mod tests;

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CommandRequest {
    pub unit_id: String,
    pub base_revision: u64,
    pub operation_id: String,
    pub command_id: String,
    pub params: Value,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct Mutation {
    pub id: String,
    pub sheet_id: String,
    pub params: Value,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum AccessRole {
    Viewer,
    Commenter,
    Editor,
    Owner,
}

pub fn required_role(request: &CommandRequest) -> KernelResult<AccessRole> {
    let ids: Vec<&str> = if request.command_id == "operation.apply" {
        request.params["mutations"]
            .as_array()
            .ok_or_else(|| invalid("mutations is required"))?
            .iter()
            .map(|m| text(m, "id"))
            .collect::<KernelResult<_>>()?
    } else {
        vec![request.command_id.as_str()]
    };
    if ids.is_empty() {
        return Err(invalid("An operation must contain a mutation"));
    }
    for id in &ids {
        if !MUTATION_IDS.contains(id) {
            return Err(KernelError::new(
                "COMMAND_UNKNOWN",
                "No public canonical mutation exists for this identifier",
            )
            .at(*id));
        }
    }
    Ok(ids
        .into_iter()
        .map(|id| {
            if id.starts_with("sheet.protect.") {
                AccessRole::Owner
            } else if id.starts_with("comment.") || id.starts_with("note.") {
                AccessRole::Commenter
            } else {
                AccessRole::Editor
            }
        })
        .max()
        .unwrap())
}

pub(crate) struct Transaction<'a> {
    pub manifest: WorkbookManifest,
    base: &'a WorkbookPages,
    original: &'a WorkbookManifest,
    writes: BTreeMap<CellAddress, Option<Cell>>,
    touched: Vec<CellAddress>,
    pub affected: Vec<RangeRef>,
}
impl<'a> Transaction<'a> {
    pub fn sheet(&self, id: &str) -> KernelResult<&SheetManifest> {
        self.manifest
            .sheets
            .iter()
            .find(|s| s.sheet_id == id)
            .ok_or_else(|| KernelError::new("SHEET_NOT_FOUND", "Worksheet does not exist").at(id))
    }
    pub fn sheet_mut(&mut self, id: &str) -> KernelResult<&mut SheetManifest> {
        self.manifest
            .sheets
            .iter_mut()
            .find(|s| s.sheet_id == id)
            .ok_or_else(|| KernelError::new("SHEET_NOT_FOUND", "Worksheet does not exist").at(id))
    }
    pub fn read(&self, sheet: &str, row: u32, column: u32) -> KernelResult<Option<Cell>> {
        let s = self.sheet(sheet)?;
        if row >= s.row_count || column >= s.column_count {
            return Err(invalid("Cell exceeds worksheet extent"));
        }
        let address = CellAddress {
            sheet_id: sheet.into(),
            row,
            column,
        };
        if let Some(cell) = self.writes.get(&address) {
            return Ok(cell.clone());
        }
        if !self
            .original
            .sheets
            .iter()
            .any(|s| s.sheet_id == sheet && row < s.row_count && column < s.column_count)
        {
            return Ok(None);
        }
        self.base.read_cell(&address)
    }
    pub fn write(
        &mut self,
        sheet: &str,
        row: u32,
        column: u32,
        cell: Option<Cell>,
    ) -> KernelResult<()> {
        let s = self.sheet(sheet)?;
        if row >= s.row_count || column >= s.column_count {
            return Err(invalid("Write exceeds worksheet extent"));
        }
        if let Some(c) = &cell {
            c.value.validate()?;
            if c.formula.as_ref().is_some_and(|f| !f.starts_with('=')) {
                return Err(invalid("Formula must start with ="));
            }
        }
        let address = CellAddress {
            sheet_id: sheet.into(),
            row,
            column,
        };
        self.touched.push(address.clone());
        self.writes.insert(address, cell);
        Ok(())
    }
    /// Iterates the page directory, never the empty logical worksheet extent.
    pub fn cells(&self, sheet: &str) -> KernelResult<Vec<(CellAddress, Cell)>> {
        self.sheet(sheet)?;
        let mut cells = BTreeMap::new();
        let base = self.original;
        if let Some(s) = base.sheets.iter().find(|s| s.sheet_id == sheet) {
            for page in base.pages.iter().filter(|p| p.sheet_id == sheet) {
                let r = RangeRef {
                    sheet_id: sheet.into(),
                    start_row: page.page_row * PAGE_ROWS,
                    end_row: ((page.page_row + 1) * PAGE_ROWS).min(s.row_count) - 1,
                    start_column: page.page_column * PAGE_COLUMNS,
                    end_column: ((page.page_column + 1) * PAGE_COLUMNS).min(s.column_count) - 1,
                };
                self.base.read_range(&r, &mut |a, c| {
                    cells.insert(a, c);
                    Ok(())
                })?;
            }
        }
        for (a, c) in &self.writes {
            if a.sheet_id == sheet {
                if let Some(c) = c {
                    cells.insert(a.clone(), c.clone());
                } else {
                    cells.remove(a);
                }
            }
        }
        let s = self.sheet(sheet)?;
        cells.retain(|a, _| a.row < s.row_count && a.column < s.column_count);
        Ok(cells.into_iter().collect())
    }
    pub fn cells_in_range(&self, r: &RangeRef) -> KernelResult<Vec<(CellAddress, Cell)>> {
        r.validate()?;
        self.sheet(&r.sheet_id)?;
        let mut cells = BTreeMap::new();
        if let Some(s) = self
            .original
            .sheets
            .iter()
            .find(|s| s.sheet_id == r.sheet_id)
        {
            for page in self
                .original
                .pages
                .iter()
                .filter(|p| p.sheet_id == r.sheet_id)
            {
                let start_row = (page.page_row * PAGE_ROWS).max(r.start_row);
                let end_row =
                    (((page.page_row + 1) * PAGE_ROWS).min(s.row_count) - 1).min(r.end_row);
                let start_column = (page.page_column * PAGE_COLUMNS).max(r.start_column);
                let end_column = (((page.page_column + 1) * PAGE_COLUMNS).min(s.column_count) - 1)
                    .min(r.end_column);
                if start_row > end_row || start_column > end_column {
                    continue;
                }
                self.base.read_range(
                    &RangeRef {
                        sheet_id: r.sheet_id.clone(),
                        start_row,
                        end_row,
                        start_column,
                        end_column,
                    },
                    &mut |a, c| {
                        cells.insert(a, c);
                        Ok(())
                    },
                )?;
            }
        }
        for (a, c) in &self.writes {
            if r.contains(a) {
                if let Some(c) = c {
                    cells.insert(a.clone(), c.clone());
                } else {
                    cells.remove(a);
                }
            }
        }
        Ok(cells.into_iter().collect())
    }
}

impl CellReader for Transaction<'_> {
    fn revision(&self) -> u64 {
        self.manifest.revision
    }
    fn read_cell(&self, a: &CellAddress) -> KernelResult<Option<Cell>> {
        self.read(&a.sheet_id, a.row, a.column)
    }
    fn read_range(
        &self,
        r: &RangeRef,
        visitor: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()> {
        r.validate()?;
        for (a, c) in self.cells_in_range(r)? {
            visitor(a, c)?;
        }
        Ok(())
    }
}
pub fn execute(pages: &mut WorkbookPages, request: CommandRequest) -> KernelResult<ChangeSet> {
    execute_authorized(pages, request, AccessRole::Editor)
}
/// Only the Java/native trusted envelope supplies this role. Browser previews
/// have no authority to persist changes, regardless of the role they request.
pub fn execute_authorized(
    pages: &mut WorkbookPages,
    request: CommandRequest,
    role: AccessRole,
) -> KernelResult<ChangeSet> {
    if role < required_role(&request)? {
        return Err(KernelError::new(
            "FORBIDDEN",
            "Workbook role does not permit this command",
        ));
    }
    let initial = pages.manifest();
    if request.unit_id != initial.unit_id {
        return Err(KernelError::new(
            "WORKBOOK_ID_MISMATCH",
            "Command belongs to another workbook",
        ));
    }
    if request.base_revision != initial.revision {
        return Err(
            KernelError::new("STALE_REVISION", "Command base revision is stale")
                .recover("refresh-manifest"),
        );
    }
    if request.operation_id.trim().is_empty() {
        return Err(invalid("operationId is required"));
    }
    let mutations: Vec<Mutation> = if request.command_id == "operation.apply" {
        serde_json::from_value(
            request
                .params
                .get("mutations")
                .cloned()
                .ok_or_else(|| invalid("mutations is required"))?,
        )
        .map_err(|e| invalid(format!("Invalid mutation list: {e}")))?
    } else {
        let sheet = request
            .params
            .get("sheetId")
            .and_then(Value::as_str)
            .unwrap_or("")
            .to_owned();
        vec![Mutation {
            id: request.command_id.clone(),
            sheet_id: sheet,
            params: request.params.clone(),
        }]
    };
    if mutations.is_empty() {
        return Err(invalid("An operation must contain a mutation"));
    }
    let mut tx = Transaction {
        manifest: initial.clone(),
        base: pages,
        original: &initial,
        writes: BTreeMap::new(),
        touched: Vec::new(),
        affected: Vec::new(),
    };
    for m in &mutations {
        if !m.params.is_object() {
            return Err(invalid("Mutation parameters must be an object").at(&m.id));
        }
        if let Some(id) = m.params.get("sheetId").and_then(Value::as_str) {
            if id != m.sheet_id {
                return Err(invalid("Mutation worksheet identity differs from params").at(&m.id));
            }
        }
        tx.touched.clear();
        let first_range = tx.affected.len();
        if !(cells::apply(&mut tx, &m.id, &m.sheet_id, &m.params)?
            || structure::apply(&mut tx, &m.id, &m.sheet_id, &m.params)?
            || sheet_metadata::apply(&mut tx, &m.id, &m.sheet_id, &m.params)?
            || objects::apply(&mut tx, &m.id, &m.sheet_id, &m.params)?
            || review::apply(&mut tx, &m.id, &m.sheet_id, &m.params)?
            || protection::apply(&mut tx, &m.id, &m.sheet_id, &m.params)?
            || common(&mut tx, &m.id, &m.sheet_id, &m.params)?)
        {
            return Err(KernelError::new(
                "COMMAND_UNKNOWN",
                "No canonical command exists for this identifier",
            )
            .at(&m.id));
        }
        for address in &tx.touched {
            tx.affected.push(RangeRef {
                sheet_id: address.sheet_id.clone(),
                start_row: address.row,
                end_row: address.row,
                start_column: address.column,
                end_column: address.column,
            });
        }
        if role != AccessRole::Owner {
            protection::check(&tx, &initial, &m.id, &tx.affected[first_range..])?;
        }
        validation::check(&tx, &m.id, &m.params, &tx.touched)?;
    }
    // A structural shrink removes trailing pages through the manifest. Blank
    // writes to the removed extent must not resurrect those pages.
    let writes = tx
        .writes
        .into_iter()
        .filter(|(a, c)| {
            c.is_some()
                || tx.manifest.sheets.iter().any(|s| {
                    s.sheet_id == a.sheet_id && a.row < s.row_count && a.column < s.column_count
                })
        })
        .map(|(address, cell)| CellWrite { address, cell })
        .collect();
    let manifest = tx.manifest;
    let affected = tx.affected;
    let mut result = pages.apply_transaction(
        request.operation_id,
        request.base_revision,
        writes,
        manifest,
    )?;
    result.affected_ranges = affected;
    Ok(result)
}

fn common(tx: &mut Transaction, id: &str, sheet: &str, p: &Value) -> KernelResult<bool> {
    match id {
        "workbook.renamed" => {
            let name = text(p, "name")?.trim();
            if name.is_empty() || name.chars().count() > 255 {
                return Err(invalid("Workbook name is invalid"));
            }
            tx.manifest.name = name.into();
        }
        "sheet.extent.grow" | "sheet.extent.restore" => {
            let rows = integer(p, "rowCount")?;
            let cols = integer(p, "columnCount")?;
            if rows == 0 || rows > MAX_ROWS || cols == 0 || cols > MAX_COLUMNS {
                return Err(invalid("Worksheet dimensions exceed bounds"));
            }
            let old = tx.sheet(sheet)?;
            if id == "sheet.extent.grow" && (rows < old.row_count || cols < old.column_count) {
                return Err(invalid("Extent grow cannot shrink"));
            }
            if rows < old.row_count || cols < old.column_count {
                if tx
                    .cells(sheet)?
                    .iter()
                    .any(|(a, _)| a.row >= rows || a.column >= cols)
                {
                    return Err(invalid("Extent restore would discard authored cells"));
                }
            }
            let s = tx.sheet_mut(sheet)?;
            s.row_count = rows;
            s.column_count = cols;
        }
        "workbook.editing.options.set" => {
            for key in [
                "allowEditDirectly",
                "moveAfterEnter",
                "formulaAutoComplete",
                "valueAutoComplete",
            ] {
                if !p[key].is_boolean() {
                    return Err(invalid(format!("{key} must be Boolean")));
                }
            }
            if !["up", "down", "left", "right"].contains(&text(p, "enterDirection")?) {
                return Err(invalid("Invalid enter direction"));
            }
            if !p
                .get("fixedDecimalPlaces")
                .is_some_and(|v| v.is_null() || v.as_u64().is_some_and(|n| n <= 15))
            {
                return Err(invalid("Invalid fixed decimal places"));
            }
            tx.manifest
                .metadata
                .insert("editingOptions".into(), p.clone());
        }
        _ => return Ok(false),
    }
    Ok(true)
}

pub(crate) fn invalid(message: impl Into<String>) -> KernelError {
    KernelError::new("VALIDATION_ERROR", message)
}
pub(crate) fn text<'a>(p: &'a Value, key: &str) -> KernelResult<&'a str> {
    p.get(key)
        .and_then(Value::as_str)
        .filter(|s| !s.trim().is_empty())
        .ok_or_else(|| invalid(format!("{key} must be nonempty text")))
}
pub(crate) fn integer(p: &Value, key: &str) -> KernelResult<u32> {
    p.get(key)
        .and_then(Value::as_u64)
        .and_then(|n| u32::try_from(n).ok())
        .ok_or_else(|| invalid(format!("{key} must be an unsigned integer")))
}
pub(crate) fn range(tx: &Transaction, sheet: &str, v: &Value) -> KernelResult<RangeRef> {
    let r: RangeRef =
        serde_json::from_value(v.clone()).map_err(|e| invalid(format!("Invalid range: {e}")))?;
    r.validate()?;
    if r.sheet_id != sheet {
        return Err(invalid("Range belongs to another worksheet"));
    }
    let s = tx.sheet(sheet)?;
    if r.end_row >= s.row_count || r.end_column >= s.column_count {
        return Err(invalid("Range exceeds worksheet dimensions"));
    }
    Ok(r)
}
