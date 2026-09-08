//! Native, file-backed OOXML ownership. XML payloads stream into canonical pages.
mod archive;
mod artifact;
mod document;
mod shared_strings;
mod styles;
mod worksheet;
mod xml;
mod xmlnode;

pub use artifact::{Artifact,ArtifactIdentity,CodecRevision};
pub use document::{CellRecord,DateSystem,DocumentFormat,FeatureCapability,FormatSupport,Metadata,NativeDocument,ResourceLimits,SheetMetadata};
pub use kernel_core::{KernelError,KernelResult};
pub const NATIVE_DOCUMENT_CODEC_REVISION:CodecRevision=CodecRevision(1);
