mod dispatch;
#[cfg(not(target_arch = "wasm32"))]
mod native_document;
#[cfg(target_arch = "wasm32")]
mod wasm;
pub use dispatch::{KernelHost, MAX_FRAME_BYTES};
