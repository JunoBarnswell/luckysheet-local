use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, thiserror::Error)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
#[error("{code}: {message}")]
pub struct KernelError {
    pub code: String,
    pub message: String,
    pub object: Option<String>,
    pub recovery: String,
}

impl KernelError {
    pub fn new(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            code: code.into(),
            message: message.into(),
            object: None,
            recovery: "correct-input".into(),
        }
    }
    pub fn at(mut self, object: impl Into<String>) -> Self {
        self.object = Some(object.into());
        self
    }
    pub fn recover(mut self, action: impl Into<String>) -> Self {
        self.recovery = action.into();
        self
    }
}
pub type KernelResult<T> = Result<T, KernelError>;
