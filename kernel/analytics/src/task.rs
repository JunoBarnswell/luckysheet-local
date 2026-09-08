use crate::source::SourceCache;
use crate::{
    filter::{FilterRequest, execute_filter},
    pivot::{PivotRequest, execute_pivot},
    query::{QueryRequest, execute_query},
};
use kernel_core::{CellReader, KernelError, KernelResult, TaskBudget};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::cell::RefCell;
use std::sync::atomic::{AtomicBool, Ordering};
#[cfg(not(target_arch = "wasm32"))]
use std::time::Instant;
#[cfg(target_arch = "wasm32")]
#[link(wasm_import_module = "env")]
unsafe extern "C" {
    fn analytics_now_ms() -> f64;
}

/// Trusted upper bounds for one analytics task. Request budgets are caller
/// hints; they can reduce these limits but can never increase them.
pub const MAX_MEMORY_BYTES: u64 = 256 * 1024 * 1024;
pub const MAX_TIMEOUT_MS: u64 = 60_000;
pub const MAX_TEMPORARY_BYTES: u64 = 16 * 1024 * 1024 * 1024;

fn bounded_budget(requested: TaskBudget) -> TaskBudget {
    TaskBudget {
        memory_bytes: requested.memory_bytes.min(MAX_MEMORY_BYTES),
        timeout_ms: requested.timeout_ms.min(MAX_TIMEOUT_MS),
        temporary_bytes: requested.temporary_bytes.min(MAX_TEMPORARY_BYTES),
    }
}
pub(crate) struct TaskClock {
    #[cfg(not(target_arch = "wasm32"))]
    instant: Instant,
    #[cfg(target_arch = "wasm32")]
    milliseconds: f64,
}
impl TaskClock {
    pub fn now() -> Self {
        Self {
            #[cfg(not(target_arch = "wasm32"))]
            instant: Instant::now(),
            #[cfg(target_arch = "wasm32")]
            milliseconds: unsafe { analytics_now_ms() },
        }
    }
    fn elapsed_ms(&self) -> f64 {
        #[cfg(not(target_arch = "wasm32"))]
        {
            self.instant.elapsed().as_secs_f64() * 1000.
        }
        #[cfg(target_arch = "wasm32")]
        {
            unsafe { analytics_now_ms() - self.milliseconds }
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum AnalyticsRequest {
    Filter(FilterRequest),
    Query(QueryRequest),
    Pivot(PivotRequest),
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum AnalyticsResponse {
    Filter(crate::FilterResult),
    Query(crate::QueryResult),
    Pivot(crate::PivotResult),
}

pub(crate) struct TaskContext<'a> {
    pub reader: &'a dyn CellReader,
    pub revision: u64,
    pub budget: TaskBudget,
    pub started: TaskClock,
    pub cancel: &'a AtomicBool,
    pub source_cache: Option<&'a RefCell<SourceCache>>,
}

/// A host creates one runtime per workbook identity and retains it between
/// requests. Source caches survive cancellation and layout/filter changes.
#[derive(Default)]
pub struct AnalyticsRuntime {
    source_cache: RefCell<SourceCache>,
}
impl AnalyticsRuntime {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn execute(
        &self,
        request: Value,
        reader: &dyn CellReader,
        cancel: &AtomicBool,
    ) -> KernelResult<Value> {
        execute_internal(request, reader, cancel, Some(&self.source_cache))
    }
    pub fn clear(&self) {
        self.source_cache.borrow_mut().clear();
    }
}

impl TaskContext<'_> {
    pub fn checkpoint(&self, phase: &str) -> KernelResult<()> {
        if self.cancel.load(Ordering::Acquire) {
            return Err(KernelError::new(
                "ANALYTICS_TASK_CANCELLED",
                "Analytics task was cancelled",
            )
            .at(phase)
            .recover("retry"));
        }
        if self.started.elapsed_ms() > self.budget.timeout_ms as f64 {
            return Err(KernelError::new(
                "ANALYTICS_TASK_TIMEOUT",
                "Analytics task exceeded its time budget",
            )
            .at(phase)
            .recover("change-query-or-budget"));
        }
        if self.reader.revision() != self.revision {
            return Err(KernelError::new(
                "ANALYTICS_REVISION_MISMATCH",
                "Pinned source revision changed during execution",
            )
            .recover("retry-on-current-revision"));
        }
        Ok(())
    }
    pub fn memory(&self, bytes: u64, phase: &str) -> KernelResult<()> {
        if bytes > self.budget.memory_bytes {
            return Err(KernelError::new(
                "ANALYTICS_MEMORY_BUDGET_EXCEEDED",
                format!(
                    "{bytes} bytes exceed the {} byte working memory budget",
                    self.budget.memory_bytes
                ),
            )
            .at(phase)
            .recover("increase-budget-or-reduce-query"));
        }
        self.checkpoint(phase)
    }
}

/// Executes one canonical request. `request` must contain `revision` and the
/// operation is rejected before reading any cells when it is absent/mismatched.
pub fn execute(request: Value, reader: &dyn CellReader) -> KernelResult<Value> {
    let cancel = AtomicBool::new(false);
    execute_with_cancel(request, reader, &cancel)
}

/// Same transaction as `execute`, with a host-owned cancellation flag. The
/// flag is sampled at every scan, aggregate, and projection checkpoint.
pub fn execute_with_cancel(
    request: Value,
    reader: &dyn CellReader,
    cancel: &AtomicBool,
) -> KernelResult<Value> {
    execute_internal(request, reader, cancel, None)
}
fn execute_internal(
    request: Value,
    reader: &dyn CellReader,
    cancel: &AtomicBool,
    source_cache: Option<&RefCell<SourceCache>>,
) -> KernelResult<Value> {
    let revision = request
        .get("revision")
        .and_then(Value::as_u64)
        .ok_or_else(|| {
            KernelError::new(
                "ANALYTICS_REVISION_REQUIRED",
                "Analytics requests require a pinned revision",
            )
        })?;
    if reader.revision() != revision {
        return Err(KernelError::new(
            "ANALYTICS_REVISION_MISMATCH",
            "Request revision is not the reader revision",
        )
        .recover("retry-on-current-revision"));
    }
    let typed: AnalyticsRequest = serde_json::from_value(request).map_err(|e| {
        KernelError::new("ANALYTICS_REQUEST_INVALID", e.to_string()).recover("correct-input")
    })?;
    let (revision, budget) = request_control(&typed)?;
    if reader.revision() != revision {
        return Err(KernelError::new(
            "ANALYTICS_REVISION_MISMATCH",
            "Request revision is not the reader revision",
        )
        .recover("retry-on-current-revision"));
    }
    let ctx = TaskContext {
        reader,
        revision,
        budget,
        started: TaskClock::now(),
        cancel,
        source_cache,
    };
    if ctx.budget.memory_bytes == 0 || ctx.budget.timeout_ms == 0 {
        return Err(KernelError::new(
            "ANALYTICS_BUDGET_INVALID",
            "Memory and timeout budgets must be positive",
        ));
    }
    ctx.checkpoint("analytics-start")?;
    let response = match typed {
        AnalyticsRequest::Filter(r) => AnalyticsResponse::Filter(execute_filter(&ctx, r)?),
        AnalyticsRequest::Query(r) => AnalyticsResponse::Query(execute_query(&ctx, r)?),
        AnalyticsRequest::Pivot(r) => AnalyticsResponse::Pivot(execute_pivot(&ctx, r)?),
    };
    ctx.checkpoint("analytics-publication")?;
    serde_json::to_value(response)
        .map_err(|e| KernelError::new("ANALYTICS_RESPONSE_INVALID", e.to_string()))
}

fn request_control(request: &AnalyticsRequest) -> KernelResult<(u64, TaskBudget)> {
    let value = serde_json::to_value(request)
        .map_err(|e| KernelError::new("ANALYTICS_REQUEST_INVALID", e.to_string()))?;
    let revision = value
        .get("revision")
        .and_then(Value::as_u64)
        .ok_or_else(|| {
            KernelError::new(
                "ANALYTICS_REVISION_REQUIRED",
                "Analytics requests require a pinned revision",
            )
        })?;
    let budget = match value.get("budget") {
        Some(v) if !v.is_null() => serde_json::from_value(v.clone())
            .map_err(|e| KernelError::new("ANALYTICS_BUDGET_INVALID", e.to_string()))?,
        _ => TaskBudget::default(),
    };
    Ok((revision, bounded_budget(budget)))
}

#[cfg(test)]
mod tests {
    use super::{MAX_MEMORY_BYTES, MAX_TEMPORARY_BYTES, MAX_TIMEOUT_MS, bounded_budget, execute};
    use kernel_core::{Cell, CellAddress, CellReader, KernelResult, RangeRef, TaskBudget};
    use serde_json::json;

    struct EmptyReader;
    impl CellReader for EmptyReader {
        fn revision(&self) -> u64 {
            7
        }
        fn read_cell(&self, _: &CellAddress) -> KernelResult<Option<Cell>> {
            Ok(None)
        }
        fn read_range(
            &self,
            _: &RangeRef,
            _: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
        ) -> KernelResult<()> {
            Ok(())
        }
    }

    #[test]
    fn rejects_missing_pinned_revision_before_scan() {
        let error = execute(json!({"kind":"query","range":{"sheetId":"s","startRow":0,"endRow":0,"startColumn":0,"endColumn":0}}), &EmptyReader).unwrap_err();
        assert_eq!(error.code, "ANALYTICS_REVISION_REQUIRED");
    }

    #[test]
    fn rejects_revision_mismatch_before_scan() {
        let error = execute(json!({"kind":"query","revision":8,"range":{"sheetId":"s","startRow":0,"endRow":0,"startColumn":0,"endColumn":0}}), &EmptyReader).unwrap_err();
        assert_eq!(error.code, "ANALYTICS_REVISION_MISMATCH");
    }

    #[test]
    fn caller_budget_is_intersected_with_trusted_limits() {
        let bounded = bounded_budget(TaskBudget {
            memory_bytes: u64::MAX,
            timeout_ms: u64::MAX,
            temporary_bytes: u64::MAX,
        });
        assert_eq!(bounded.memory_bytes, MAX_MEMORY_BYTES);
        assert_eq!(bounded.timeout_ms, MAX_TIMEOUT_MS);
        assert_eq!(bounded.temporary_bytes, MAX_TEMPORARY_BYTES);
    }

    #[test]
    fn budget_within_trusted_limits_is_preserved() {
        let requested = TaskBudget {
            memory_bytes: 8 * 1024 * 1024,
            timeout_ms: 2_000,
            temporary_bytes: 32 * 1024 * 1024,
        };
        let bounded = bounded_budget(requested.clone());
        assert_eq!(bounded.memory_bytes, requested.memory_bytes);
        assert_eq!(bounded.timeout_ms, requested.timeout_ms);
        assert_eq!(bounded.temporary_bytes, requested.temporary_bytes);
    }
}
