use crate::{KernelError, KernelResult};
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;

pub const KERNEL_PROTOCOL_VERSION: u32 = 1;
pub const WORKBOOK_MANIFEST_VERSION: u32 = 11;
pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLUMNS: u32 = 16_384;
pub const PAGE_ROWS: u32 = 1024;
pub const PAGE_COLUMNS: u32 = 32;

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(untagged)]
pub enum Scalar {
    Null,
    Boolean(bool),
    Number(f64),
    Text(String),
    Error(FormulaError),
}
impl Default for Scalar {
    fn default() -> Self {
        Self::Null
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FormulaError {
    pub kind: String,
    pub code: String,
    pub message: String,
}
impl Scalar {
    pub fn error(code: &str, message: impl Into<String>) -> Self {
        Self::Error(FormulaError {
            kind: "error".into(),
            code: code.into(),
            message: message.into(),
        })
    }
    pub fn validate(&self) -> KernelResult<()> {
        match self {
            Self::Number(n) if !n.is_finite() => Err(KernelError::new(
                "CELL_VALUE_INVALID",
                "Numbers must be finite",
            )),
            Self::Error(e)
                if e.kind != "error"
                    || ![
                        "#NULL!",
                        "#DIV/0!",
                        "#VALUE!",
                        "#REF!",
                        "#NAME?",
                        "#NUM!",
                        "#N/A",
                        "#CALC!",
                        "#SPILL!",
                        "#BLOCKED!",
                        "#GETTING_DATA",
                    ]
                    .contains(&e.code.as_str()) =>
            {
                Err(KernelError::new(
                    "CELL_VALUE_INVALID",
                    "Unknown formula error",
                ))
            }
            _ => Ok(()),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Cell {
    pub value: Scalar,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula: Option<String>,
    /// Authored metadata is retained exactly; interpretation belongs to the corresponding kernel domain.
    #[serde(flatten)]
    pub metadata: BTreeMap<String, serde_json::Value>,
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct CellAddress {
    pub sheet_id: String,
    pub row: u32,
    pub column: u32,
}
impl CellAddress {
    pub fn validate(&self) -> KernelResult<()> {
        if self.sheet_id.is_empty() || self.row >= MAX_ROWS || self.column >= MAX_COLUMNS {
            return Err(KernelError::new(
                "CELL_ADDRESS_INVALID",
                "Cell address is outside the worksheet",
            )
            .at(format!("{}:{}:{}", self.sheet_id, self.row, self.column)));
        }
        Ok(())
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct RangeRef {
    pub sheet_id: String,
    pub start_row: u32,
    pub end_row: u32,
    pub start_column: u32,
    pub end_column: u32,
}
impl RangeRef {
    pub fn validate(&self) -> KernelResult<()> {
        if self.sheet_id.is_empty()
            || self.start_row > self.end_row
            || self.start_column > self.end_column
            || self.end_row >= MAX_ROWS
            || self.end_column >= MAX_COLUMNS
        {
            return Err(KernelError::new("RANGE_INVALID", "Invalid worksheet range")
                .at(self.sheet_id.clone()));
        }
        Ok(())
    }
    pub fn contains(&self, address: &CellAddress) -> bool {
        self.sheet_id == address.sheet_id
            && address.row >= self.start_row
            && address.row <= self.end_row
            && address.column >= self.start_column
            && address.column <= self.end_column
    }
    pub fn intersects(&self, other: &Self) -> bool {
        self.sheet_id == other.sheet_id
            && self.start_row <= other.end_row
            && self.end_row >= other.start_row
            && self.start_column <= other.end_column
            && self.end_column >= other.start_column
    }
}

/// Read-only revision-pinned input shared by all computational domains. Missing data pages are errors, never blank values.
pub trait CellReader {
    fn revision(&self) -> u64;
    fn read_cell(&self, address: &CellAddress) -> KernelResult<Option<Cell>>;
    fn read_range(
        &self,
        range: &RangeRef,
        visitor: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
    ) -> KernelResult<()>;
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TaskBudget {
    pub memory_bytes: u64,
    pub timeout_ms: u64,
    pub temporary_bytes: u64,
}
impl Default for TaskBudget {
    fn default() -> Self {
        Self {
            memory_bytes: 2 * 1024 * 1024 * 1024,
            timeout_ms: 60_000,
            temporary_bytes: 16 * 1024 * 1024 * 1024,
        }
    }
}
