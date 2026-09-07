use kernel_core::{KernelError, KernelResult};
use std::collections::BTreeMap;
use std::fs::{self, File};
use std::io::{self, BufRead, BufReader, Read};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use zip::{write::SimpleFileOptions, ZipArchive, ZipWriter};

/// A bounded OPC package view. Parts are opened and decompressed on demand.
#[derive(Debug, Clone)]
pub(crate) struct Package {
    source: Arc<PackageSource>,
    limits: crate::ResourceLimits,
}
#[derive(Debug)]
enum PackageSource {
    File { path: PathBuf, remove_on_drop: bool },
}
impl Drop for PackageSource {
    fn drop(&mut self) {
        if let Self::File {
            path,
            remove_on_drop: true,
        } = self
        {
            let _ = fs::remove_file(path);
        }
    }
}

impl Package {
    pub(crate) fn open(
        path: impl AsRef<Path>,
        limits: crate::ResourceLimits,
    ) -> KernelResult<Self> {
        let path = path.as_ref().to_path_buf();
        let metadata = fs::metadata(&path).map_err(|e| invalid("ZIP_OPEN", e.to_string()))?;
        if metadata.len() > limits.max_archive_bytes {
            return Err(limit(
                "ZIP_BYTES",
                "archive exceeds the configured compressed input budget",
            ));
        }
        let package = Self {
            source: Arc::new(PackageSource::File {
                path,
                remove_on_drop: false,
            }),
            limits,
        };
        package.validate_index()?;
        Ok(package)
    }

    pub(crate) fn part_names(&self) -> KernelResult<Vec<String>> {
        let mut archive = self.open_archive()?;
        let mut names = Vec::with_capacity(archive.len());
        for index in 0..archive.len() {
            let entry = archive
                .by_index(index)
                .map_err(|e| invalid("ZIP_ENTRY_INVALID", e.to_string()))?;
            if entry.is_dir() {
                continue;
            }
            names.push(normalize_entry_name(entry.name())?);
        }
        names.sort();
        names.dedup();
        Ok(names)
    }

    /// Gives a bounded buffered reader and drains the member after the callback.
    pub(crate) fn with_part_reader<T>(
        &self,
        name: &str,
        callback: impl FnOnce(&mut dyn BufRead) -> KernelResult<T>,
    ) -> KernelResult<T> {
        let canonical = normalize_requested_name(name)?;
        let mut archive = self.open_archive()?;
        let mut entry = archive
            .by_name(&canonical)
            .map_err(|e| invalid("ZIP_PART_MISSING", e.to_string()))?;
        let declared = entry.size();
        if declared > self.limits.max_entry_bytes {
            return Err(limit(
                "ZIP_ENTRY_BYTES",
                "archive part exceeds the configured per-part budget",
            ));
        }
        if self.limits.max_temporary_bytes == 0 {
            return Err(limit(
                "ZIP_TEMPORARY_BYTES",
                "streaming reader requires a positive temporary buffer budget",
            ));
        }
        let buffer_size = usize::try_from(self.limits.max_temporary_bytes.min(128 * 1024).max(1))
            .map_err(|_| {
            limit(
                "ZIP_TEMPORARY_BYTES",
                "temporary reader buffer size is not representable",
            )
        })?;
        let mut bounded = BoundedEntry::new(&mut entry, self.limits.max_entry_bytes);
        let mut reader = BufReader::with_capacity(buffer_size, &mut bounded);
        let callback_result = callback(&mut reader)?;
        let mut sink = io::sink();
        io::copy(&mut reader, &mut sink).map_err(|e| invalid("ZIP_READ", e.to_string()))?;
        drop(reader);
        if bounded.observed != declared {
            return Err(invalid(
                "ZIP_ENTRY_SIZE_MISMATCH",
                format!(
                    "part {canonical} declared {declared} bytes but yielded {}",
                    bounded.observed
                ),
            ));
        }
        Ok(callback_result)
    }

    /// Rewrites to a file while copying untouched members directly from source.
    /// Replacements must name existing parts.
    pub(crate) fn write_to_path(
        &self,
        output: impl AsRef<Path>,
        replacements: &BTreeMap<String, PathBuf>,
    ) -> KernelResult<()> {
        let output = output.as_ref();
        let file = File::create(output).map_err(|e| invalid("ZIP_WRITE", e.to_string()))?;
        let mut writer = ZipWriter::new(file);
        let options =
            SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
        let mut archive = self.open_archive()?;
        let mut seen = std::collections::BTreeSet::new();
        for index in 0..archive.len() {
            let mut entry = archive
                .by_index(index)
                .map_err(|e| invalid("ZIP_ENTRY_INVALID", e.to_string()))?;
            if entry.is_dir() {
                continue;
            }
            let name = normalize_entry_name(entry.name())?;
            if !seen.insert(name.clone()) {
                return Err(invalid("ZIP_DUPLICATE_PART", name));
            }
            writer
                .start_file(&name, options)
                .map_err(|e| invalid("ZIP_WRITE", e.to_string()))?;
            if let Some(replacement) = replacements.get(&name) {
                if fs::metadata(replacement)
                    .map_err(|e| invalid("ZIP_REPLACEMENT_READ", e.to_string()))?
                    .len()
                    > self.limits.max_entry_bytes
                {
                    return Err(limit(
                        "ZIP_ENTRY_BYTES",
                        "replacement part exceeds the configured per-part budget",
                    ));
                }
                let mut replacement_file = File::open(replacement)
                    .map_err(|e| invalid("ZIP_REPLACEMENT_READ", e.to_string()))?;
                io::copy(&mut replacement_file, &mut writer)
                    .map_err(|e| invalid("ZIP_WRITE", e.to_string()))?;
            } else {
                let declared = entry.size();
                let mut bounded = BoundedEntry::new(&mut entry, self.limits.max_entry_bytes);
                io::copy(&mut bounded, &mut writer)
                    .map_err(|e| invalid("ZIP_WRITE", e.to_string()))?;
                if bounded.observed != declared {
                    return Err(invalid(
                        "ZIP_ENTRY_SIZE_MISMATCH",
                        format!(
                            "part {name} declared {declared} bytes but yielded {}",
                            bounded.observed
                        ),
                    ));
                }
            }
        }
        for name in replacements.keys() {
            if !seen.contains(name) {
                return Err(invalid("ZIP_REPLACEMENT_MISSING", name.clone()));
            }
        }
        let file = writer
            .finish()
            .map_err(|e| invalid("ZIP_WRITE", e.to_string()))?;
        file.sync_all()
            .map_err(|e| invalid("ZIP_WRITE", e.to_string()))?;
        if file
            .metadata()
            .map_err(|e| invalid("ZIP_WRITE", e.to_string()))?
            .len()
            > self.limits.max_archive_bytes
        {
            return Err(limit(
                "ZIP_BYTES",
                "output exceeds the configured compressed archive budget",
            ));
        }
        Ok(())
    }

    fn open_archive(&self) -> KernelResult<ZipArchive<File>> {
        let path = match self.source.as_ref() {
            PackageSource::File { path, .. } => path,
        };
        let file = File::open(path).map_err(|e| invalid("ZIP_OPEN", e.to_string()))?;
        ZipArchive::new(file).map_err(|e| invalid("ZIP_INVALID", e.to_string()))
    }
    fn validate_index(&self) -> KernelResult<()> {
        let mut archive = self.open_archive()?;
        if archive.len() as u64 > self.limits.max_entries {
            return Err(limit("ZIP_ENTRIES", "archive has too many parts"));
        }
        let mut expanded = 0u64;
        let mut names = std::collections::BTreeSet::new();
        for index in 0..archive.len() {
            let entry = archive
                .by_index(index)
                .map_err(|e| invalid("ZIP_ENTRY_INVALID", e.to_string()))?;
            if entry.is_dir() {
                continue;
            }
            let name = normalize_entry_name(entry.name())?;
            if !names.insert(name) {
                return Err(invalid("ZIP_DUPLICATE_PART", entry.name()));
            }
            let size = entry.size();
            if size > self.limits.max_entry_bytes {
                return Err(limit(
                    "ZIP_ENTRY_BYTES",
                    "archive part exceeds the configured per-part budget",
                ));
            }
            expanded = expanded
                .checked_add(size)
                .ok_or_else(|| limit("ZIP_EXPANDED_BYTES", "expanded archive size overflow"))?;
            if expanded > self.limits.max_uncompressed_bytes {
                return Err(limit(
                    "ZIP_EXPANDED_BYTES",
                    "archive exceeds the configured expanded budget",
                ));
            }
        }
        Ok(())
    }
}

struct BoundedEntry<'a, R> {
    source: &'a mut R,
    observed: u64,
    max: u64,
}
impl<'a, R> BoundedEntry<'a, R> {
    fn new(source: &'a mut R, max: u64) -> Self {
        Self {
            source,
            observed: 0,
            max,
        }
    }
}
impl<R: Read> Read for BoundedEntry<'_, R> {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        if self.observed >= self.max {
            return Ok(0);
        }
        let remaining = self.max - self.observed;
        let request = buffer
            .len()
            .min(usize::try_from(remaining).unwrap_or(usize::MAX));
        if request == 0 {
            return Ok(0);
        }
        let count = self.source.read(&mut buffer[..request])?;
        self.observed = self.observed.saturating_add(count as u64);
        Ok(count)
    }
}
fn normalize_requested_name(name: &str) -> KernelResult<String> {
    let canonical = normalize_entry_name(name)?;
    if canonical != name {
        return Err(invalid("ZIP_PATH_INVALID", name));
    }
    Ok(canonical)
}
fn normalize_entry_name(name: &str) -> KernelResult<String> {
    if name.is_empty() || name.starts_with('/') || name.contains('\\') {
        return Err(invalid("ZIP_PATH_INVALID", name));
    }
    let mut parts = Vec::new();
    for part in name.split('/') {
        if part.is_empty() || part == "." || part == ".." {
            return Err(invalid("ZIP_PATH_INVALID", name));
        }
        parts.push(part);
    }
    Ok(parts.join("/"))
}
fn invalid(code: &str, message: impl Into<String>) -> KernelError {
    KernelError::new(code, message).recover("Provide a valid OPC package")
}
fn limit(code: &str, message: impl Into<String>) -> KernelError {
    KernelError::new(code, message)
        .recover("Reduce the document size or raise the explicit resource budget")
}
