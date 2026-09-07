//! Canonical, revision-pinned analytics execution for the spreadsheet kernel.
//! The crate deliberately owns no workbook state: every operation is a pure
//! read against `CellReader`, and failed operations cannot publish partial state.

mod aggregate;
mod filter;
mod source;
mod pivot;
mod pivot_expression;
mod pivot_group;
mod query;
mod task;
#[cfg(test)]
mod integration_tests;

pub use filter::{
    DomainValue, FilterColumn, FilterOperator, FilterJoin, FilterOwner, FilterPredicate, FilterRequest, FilterResult, VisibilityBitmap,
    VisibilityReason,
};
pub use pivot::{AggregateKind, PivotRequest, PivotResult};
pub use pivot_expression::{CalculatedField, CalculatedItem, Evaluator, FieldDescriptor};
pub use pivot_group::{
    DateGroupUnit, Group, ManualGroup, MemberKey, MemberType, MemberValue, group, typed_member,
};
pub use query::{QueryRequest, QueryResult};
pub use task::{AnalyticsRequest, AnalyticsResponse, AnalyticsRuntime, execute, execute_with_cancel};
