use crate::{aggregate::scalar_key, task::TaskContext};
use kernel_core::{KernelError, KernelResult, RangeRef, Scalar};
use std::collections::HashMap;
use std::sync::Arc;

#[derive(Default)]
pub(crate) struct SourceCache {
    revision: Option<u64>,
    entries: HashMap<String, Arc<SourceIndex>>,
    bytes: u64,
}
impl SourceCache {
    pub fn clear(&mut self) {
        self.entries.clear();
        self.bytes = 0;
        self.revision = None;
    }
}

/// One revision-owned columnar source. Scalar objects exist only in dictionaries;
/// numeric and logical columns have contiguous vectors and validity bitmaps.
#[derive(Debug, Clone)]
pub struct SourceIndex {
    pub revision: u64,
    pub range: RangeRef,
    pub row_count: usize,
    pub columns: Vec<u32>,
    pub vectors: Vec<SourceColumn>,
    pub bytes: u64,
    ordinals: HashMap<u32, usize>,
}
#[derive(Debug, Clone)]
pub enum SourceColumn {
    Empty {
        rows: usize,
    },
    Number {
        values: Vec<f64>,
        validity: Vec<u64>,
    },
    Boolean {
        values: Vec<u8>,
        validity: Vec<u64>,
    },
    Dictionary {
        dictionary: Vec<Scalar>,
        codes: Vec<u32>,
    },
}
struct ColumnBuilder {
    vector: SourceColumn,
    dictionary_index: HashMap<String, u32>,
    bytes: u64,
}
impl SourceIndex {
    pub(crate) fn build(
        ctx: &TaskContext<'_>,
        range: &RangeRef,
        columns: &[u32],
    ) -> KernelResult<Arc<Self>> {
        let key = format!(
            "{}:{}:{}:{}:{}:{columns:?}",
            range.sheet_id, range.start_row, range.end_row, range.start_column, range.end_column
        );
        if let Some(cache) = ctx.source_cache {
            let mut cache = cache.borrow_mut();
            if cache.revision != Some(ctx.revision) {
                cache.clear();
                cache.revision = Some(ctx.revision);
            }
            if let Some(source) = cache.entries.get(&key) {
                if source.revision != ctx.revision {
                    return Err(KernelError::new(
                        "ANALYTICS_REVISION_MISMATCH",
                        "Source cache entry revision does not match pinned request",
                    ));
                }
                ctx.memory(source.bytes, "source-cache-hit")?;
                return Ok(source.clone());
            }
        }
        let source = Arc::new(Self::scan(ctx, range, columns)?);
        if let Some(cache) = ctx.source_cache {
            let mut cache = cache.borrow_mut();
            if cache.bytes + source.bytes > ctx.budget.memory_bytes / 2 {
                cache.clear();
                cache.revision = Some(ctx.revision);
            }
            if source.bytes <= ctx.budget.memory_bytes / 2 {
                cache.bytes += source.bytes;
                cache.entries.insert(key, source.clone());
            }
        }
        Ok(source)
    }
    fn scan(ctx: &TaskContext<'_>, range: &RangeRef, columns: &[u32]) -> KernelResult<Self> {
        range.validate()?;
        let selected = if columns.is_empty() {
            (range.start_column..=range.end_column).collect::<Vec<_>>()
        } else {
            columns.to_vec()
        };
        let row_count = (range.end_row - range.start_row + 1) as usize;
        let mut ordinals = HashMap::new();
        for (ordinal, column) in selected.iter().enumerate() {
            if *column < range.start_column || *column > range.end_column {
                return Err(KernelError::new(
                    "ANALYTICS_COLUMN_OUTSIDE_RANGE",
                    "Projected column is outside source range",
                )
                .at(column.to_string()));
            }
            if ordinals.insert(*column, ordinal).is_some() {
                return Err(KernelError::new(
                    "ANALYTICS_COLUMN_DUPLICATE",
                    "Source projection columns must be unique",
                ));
            }
        }
        // Check before vector allocation; heterogeneous dictionary growth is
        // charged incrementally as cells arrive from the page reader.
        let minimum = (row_count as u64)
            .checked_mul(selected.len() as u64)
            .and_then(|n| n.checked_mul(4))
            .ok_or_else(|| {
                KernelError::new(
                    "ANALYTICS_MEMORY_BUDGET_EXCEEDED",
                    "Source allocation size overflow",
                )
            })?;
        ctx.memory(minimum, "source-allocation")?;
        let mut builders = selected
            .iter()
            .map(|_| ColumnBuilder {
                vector: SourceColumn::Empty { rows: row_count },
                dictionary_index: HashMap::new(),
                bytes: 0,
            })
            .collect::<Vec<_>>();
        let mut bytes = 0;
        let mut scanned = 0;
        ctx.reader.read_range(range, &mut |address, cell| {
            if address.sheet_id != range.sheet_id
                || address.row < range.start_row
                || address.row > range.end_row
                || address.column < range.start_column
                || address.column > range.end_column
            {
                return Err(KernelError::new(
                    "ANALYTICS_SOURCE_ADDRESS_INVALID",
                    "Range reader returned a cell outside the pinned range",
                ));
            }
            if let Some(ordinal) = ordinals.get(&address.column) {
                let builder = &mut builders[*ordinal];
                let before = builder.bytes;
                builder.push((address.row - range.start_row) as usize, cell.value)?;
                bytes = bytes + builder.bytes - before;
            }
            scanned += 1;
            if scanned % 256 == 0 {
                ctx.memory(bytes, "source-scan")?;
            }
            Ok(())
        })?;
        ctx.memory(bytes, "source-publication")?;
        Ok(Self {
            revision: ctx.revision,
            range: range.clone(),
            row_count,
            columns: selected,
            vectors: builders.into_iter().map(|builder| builder.vector).collect(),
            bytes,
            ordinals,
        })
    }
    pub(crate) fn ordinal(&self, column: u32) -> KernelResult<usize> {
        self.ordinals.get(&column).copied().ok_or_else(|| {
            KernelError::new(
                "ANALYTICS_COLUMN_UNAVAILABLE",
                "Column is not present in source projection",
            )
            .at(column.to_string())
        })
    }
    pub(crate) fn column(&self, column: u32) -> KernelResult<&SourceColumn> {
        Ok(&self.vectors[self.ordinal(column)?])
    }
    pub(crate) fn value_at(&self, column: u32, row: usize) -> KernelResult<Scalar> {
        if row >= self.row_count {
            return Err(KernelError::new(
                "ANALYTICS_ROW_OUT_OF_BOUNDS",
                "Source row is outside index",
            ));
        }
        Ok(self.column(column)?.value_at(row))
    }
}
impl ColumnBuilder {
    fn push(&mut self, row: usize, value: Scalar) -> KernelResult<()> {
        value.validate()?;
        if matches!(value, Scalar::Null) {
            return Ok(());
        }
        if let SourceColumn::Empty { rows } = self.vector {
            self.vector = match &value {
                Scalar::Number(_) => {
                    self.bytes = rows as u64 * 8 + rows.div_ceil(64) as u64 * 8;
                    SourceColumn::Number {
                        values: vec![0.; rows],
                        validity: vec![0; rows.div_ceil(64)],
                    }
                }
                Scalar::Boolean(_) => {
                    self.bytes = rows as u64 + rows.div_ceil(64) as u64 * 8;
                    SourceColumn::Boolean {
                        values: vec![0; rows],
                        validity: vec![0; rows.div_ceil(64)],
                    }
                }
                _ => {
                    self.bytes = rows as u64 * 4;
                    SourceColumn::Dictionary {
                        dictionary: vec![],
                        codes: vec![0; rows],
                    }
                }
            };
        }
        match (&mut self.vector, &value) {
            (SourceColumn::Number { values, validity }, Scalar::Number(number)) => {
                values[row] = *number;
                validity[row / 64] |= 1 << (row % 64);
                return Ok(());
            }
            (SourceColumn::Boolean { values, validity }, Scalar::Boolean(boolean)) => {
                values[row] = u8::from(*boolean);
                validity[row / 64] |= 1 << (row % 64);
                return Ok(());
            }
            (SourceColumn::Dictionary { .. }, _) => {}
            _ => self.promote()?,
        }
        self.insert(row, value)
    }
    fn promote(&mut self) -> KernelResult<()> {
        let old = std::mem::replace(&mut self.vector, SourceColumn::Empty { rows: 0 });
        let rows = match &old {
            SourceColumn::Number { values, .. } => values.len(),
            SourceColumn::Boolean { values, .. } => values.len(),
            _ => {
                return Err(KernelError::new(
                    "ANALYTICS_COLUMN_STATE_INVALID",
                    "Only typed vectors can be promoted to dictionary storage",
                ));
            }
        };
        self.vector = SourceColumn::Dictionary {
            dictionary: vec![],
            codes: vec![0; rows],
        };
        // Peak old-vector storage remains charged until promotion finishes.
        self.bytes += rows as u64 * 4;
        for row in 0..rows {
            let value = old.value_at(row);
            if !matches!(value, Scalar::Null) {
                self.insert(row, value)?;
            }
        }
        Ok(())
    }
    fn insert(&mut self, row: usize, value: Scalar) -> KernelResult<()> {
        let key = scalar_key(&value);
        let SourceColumn::Dictionary { dictionary, codes } = &mut self.vector else {
            return Err(KernelError::new(
                "ANALYTICS_COLUMN_STATE_INVALID",
                "Dictionary insertion requires dictionary storage",
            ));
        };
        let slot = codes.get_mut(row).ok_or_else(|| {
            KernelError::new(
                "ANALYTICS_ROW_OUT_OF_BOUNDS",
                "Cell address is outside source range",
            )
        })?;
        let code = if let Some(id) = self.dictionary_index.get(&key) {
            *id
        } else {
            let id = u32::try_from(dictionary.len() + 1).map_err(|_| {
                KernelError::new(
                    "ANALYTICS_DICTIONARY_OVERFLOW",
                    "Dictionary cannot represent another member",
                )
            })?;
            self.bytes += std::mem::size_of::<Scalar>() as u64 + key.len() as u64 * 2 + 64;
            dictionary.push(value);
            self.dictionary_index.insert(key, id);
            id
        };
        *slot = code;
        Ok(())
    }
}
impl SourceColumn {
    pub(crate) fn value_at(&self, row: usize) -> Scalar {
        match self {
            Self::Empty { .. } => Scalar::Null,
            Self::Number { values, validity } => {
                if validity[row / 64] & (1 << (row % 64)) == 0 {
                    Scalar::Null
                } else {
                    Scalar::Number(values[row])
                }
            }
            Self::Boolean { values, validity } => {
                if validity[row / 64] & (1 << (row % 64)) == 0 {
                    Scalar::Null
                } else {
                    Scalar::Boolean(values[row] != 0)
                }
            }
            Self::Dictionary { dictionary, codes } => {
                if codes[row] == 0 {
                    Scalar::Null
                } else {
                    dictionary[codes[row] as usize - 1].clone()
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn mixed_type_promotion_preserves_every_cell() {
        let mut builder = ColumnBuilder {
            vector: SourceColumn::Empty { rows: 4 },
            dictionary_index: HashMap::new(),
            bytes: 0,
        };
        for (row, value) in [
            Scalar::Boolean(true),
            Scalar::Number(7.),
            Scalar::Text("7".into()),
            Scalar::Null,
        ]
        .into_iter()
        .enumerate()
        {
            builder.push(row, value).unwrap();
        }
        assert_eq!(builder.vector.value_at(0), Scalar::Boolean(true));
        assert_eq!(builder.vector.value_at(1), Scalar::Number(7.));
        assert_eq!(builder.vector.value_at(2), Scalar::Text("7".into()));
        assert_eq!(builder.vector.value_at(3), Scalar::Null);
    }
}
