use kernel_core::{KernelError, KernelResult};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use std::{
    fs::File,
    io::Read,
    path::{Path, PathBuf},
};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct CodecRevision(pub u32);

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ArtifactIdentity {
    pub revision: u64,
    pub checksum: String,
    pub byte_length: u64,
    pub format: String,
    pub codec_revision: CodecRevision,
}

/// Native artifacts are task-directory files; control frames never contain bytes.
#[derive(Debug, Clone)]
pub struct Artifact {
    pub identity: ArtifactIdentity,
    pub file_name: String,
    pub path: PathBuf,
}
impl Artifact {
    pub fn from_path(
        path: impl AsRef<Path>,
        revision: u64,
        format: String,
        codec_revision: CodecRevision,
    ) -> KernelResult<Self> {
        let path = path.as_ref().to_path_buf();
        let file_name = path
            .file_name()
            .and_then(|s| s.to_str())
            .ok_or_else(|| {
                KernelError::new("DOCUMENT_PATH_INVALID", "Artifact filename is invalid")
            })?
            .to_owned();
        let (checksum, byte_length) = hash_file(&path)?;
        Ok(Self {
            identity: ArtifactIdentity {
                revision,
                checksum,
                byte_length,
                format,
                codec_revision,
            },
            file_name,
            path,
        })
    }
    pub fn verify(&self) -> KernelResult<()> {
        let (checksum, byte_length) = hash_file(&self.path)?;
        if checksum != self.identity.checksum || byte_length != self.identity.byte_length {
            return Err(KernelError::new(
                "ARTIFACT_CHECKSUM_MISMATCH",
                "Native file changed after its revision was bound",
            )
            .at(self.file_name.clone())
            .recover("Reload the source artifact and retry"));
        }
        Ok(())
    }
}
fn hash_file(path: &Path) -> KernelResult<(String, u64)> {
    let mut input =
        File::open(path).map_err(|e| KernelError::new("DOCUMENT_OPEN", e.to_string()))?;
    let mut hasher = Sha256::new();
    let mut byte_length = 0u64;
    let mut buffer = [0u8; 128 * 1024];
    loop {
        let count = input
            .read(&mut buffer)
            .map_err(|e| KernelError::new("DOCUMENT_READ", e.to_string()))?;
        if count == 0 {
            break;
        }
        byte_length += count as u64;
        hasher.update(&buffer[..count]);
    }
    Ok((
        hasher
            .finalize()
            .iter()
            .map(|b| format!("{b:02x}"))
            .collect(),
        byte_length,
    ))
}
