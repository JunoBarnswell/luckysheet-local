use crate::{
    archive::Package,
    worksheet::{self, SharedString},
    xml, ResourceLimits,
};
use kernel_core::{KernelError, KernelResult};
use quick_xml::{events::Event, Reader, Writer};
use std::{
    fs::{File, OpenOptions},
    io::{BufRead, Read, Seek, SeekFrom, Write},
    path::{Path, PathBuf},
    sync::atomic::{AtomicU64, Ordering},
};

/// The shared-string index and values live on disk. Lookup retains one value,
/// regardless of the number or total text bytes of the workbook dictionary.
pub(crate) struct SharedStrings {
    index: File,
    data: File,
    count: u64,
    paths: [PathBuf; 2],
}
impl Drop for SharedStrings {
    fn drop(&mut self) {
        for path in &self.paths {
            let _ = std::fs::remove_file(path);
        }
    }
}
impl SharedStrings {
    pub fn build(
        package: &Package,
        part: Option<&str>,
        directory: &Path,
        limits: &ResourceLimits,
    ) -> KernelResult<Self> {
        static NEXT: AtomicU64 = AtomicU64::new(0);
        let n = NEXT.fetch_add(1, Ordering::Relaxed);
        let prefix = format!("native-strings-{}-{n}", std::process::id());
        let paths = [
            directory.join(format!("{prefix}.idx")),
            directory.join(format!("{prefix}.dat")),
        ];
        let index = OpenOptions::new()
            .create_new(true)
            .read(true)
            .write(true)
            .open(&paths[0])
            .map_err(io_error)?;
        let data = match OpenOptions::new()
            .create_new(true)
            .read(true)
            .write(true)
            .open(&paths[1])
        {
            Ok(f) => f,
            Err(e) => {
                drop(index);
                let _ = std::fs::remove_file(&paths[0]);
                return Err(io_error(e));
            }
        };
        let mut dictionary = Self {
            index,
            data,
            count: 0,
            paths,
        };
        if let Some(part) = part {
            package.with_part_reader(part, |input| dictionary.import(input, limits))?;
        }
        Ok(dictionary)
    }
    fn import(&mut self, input: &mut dyn BufRead, limits: &ResourceLimits) -> KernelResult<()> {
        let mut reader = Reader::from_reader(input);
        let mut buffer = Vec::new();
        let mut depth = 0usize;
        let mut item: Option<Writer<Vec<u8>>> = None;
        let mut item_depth = 0usize;
        let mut offset = 0u64;
        loop {
            let event = reader.read_event_into(&mut buffer).map_err(xml_error)?;
            match &event {
                Event::DocType(_) => {
                    return Err(KernelError::new(
                        "XML_DOCTYPE_FORBIDDEN",
                        "Shared strings cannot declare entities",
                    ))
                }
                Event::Start(e) => {
                    depth += 1;
                    if depth > limits.max_xml_depth as usize {
                        return Err(KernelError::new(
                            "XML_DEPTH_LIMIT",
                            "Shared string XML nesting exceeds budget",
                        ));
                    }
                    if xml::local_name(e.name().as_ref()) == b"si" {
                        if item.is_some() {
                            return Err(KernelError::new(
                                "OOXML_SHARED_STRINGS_INVALID",
                                "Nested si",
                            ));
                        }
                        item = Some(Writer::new(Vec::new()));
                        item_depth = depth;
                    }
                }
                Event::Empty(e) if xml::local_name(e.name().as_ref()) == b"si" => {
                    if item.is_some() {
                        return Err(KernelError::new(
                            "OOXML_SHARED_STRINGS_INVALID",
                            "Nested si",
                        ));
                    }
                    item = Some(Writer::new(Vec::new()));
                    item_depth = depth;
                }
                Event::Eof => {
                    if depth != 0 || item.is_some() {
                        return Err(KernelError::new("XML_INVALID", "Truncated shared strings"));
                    }
                    break;
                }
                _ => {}
            }
            if let Some(writer) = item.as_mut() {
                writer.write_event(event.clone()).map_err(io_error)?;
                if writer.get_ref().len() as u64 > limits.max_cell_bytes {
                    return Err(KernelError::new(
                        "CELL_TEXT_LIMIT",
                        "Shared string exceeds cell budget",
                    ));
                }
            }
            let finish = matches!(&event,Event::End(e) if xml::local_name(e.name().as_ref())==b"si"&&depth==item_depth)
                || matches!(&event,Event::Empty(e) if xml::local_name(e.name().as_ref())==b"si"&&depth==item_depth);
            if finish {
                let bytes = item.take().unwrap().into_inner();
                let value = worksheet::parse_shared_string(&bytes)?;
                let payload = serde_json::to_vec(&value).map_err(xml_error)?;
                let length = payload.len() as u64;
                if offset
                    .saturating_add(length)
                    .saturating_add((self.count + 1) * 16)
                    > limits.max_temporary_bytes
                {
                    return Err(KernelError::new(
                        "TEMPORARY_BYTES_LIMIT",
                        "Shared string disk index exceeds task budget",
                    ));
                }
                self.index
                    .write_all(&offset.to_le_bytes())
                    .map_err(io_error)?;
                self.index
                    .write_all(&length.to_le_bytes())
                    .map_err(io_error)?;
                self.data.write_all(&payload).map_err(io_error)?;
                offset += length;
                self.count += 1;
                if self.count > limits.max_cells {
                    return Err(KernelError::new(
                        "SHARED_STRINGS_LIMIT",
                        "String item count exceeds budget",
                    ));
                }
            }
            if matches!(event, Event::End(_)) {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    KernelError::new("XML_INVALID", "Unexpected shared string closing tag")
                })?;
            }
            buffer.clear();
        }
        self.index.flush().map_err(io_error)?;
        self.data.flush().map_err(io_error)?;
        Ok(())
    }
    pub fn get(&mut self, index: usize) -> KernelResult<SharedString> {
        if index as u64 >= self.count {
            return Err(KernelError::new(
                "SHARED_STRING_INDEX_INVALID",
                index.to_string(),
            ));
        }
        self.index
            .seek(SeekFrom::Start(index as u64 * 16))
            .map_err(io_error)?;
        let mut entry = [0u8; 16];
        self.index.read_exact(&mut entry).map_err(io_error)?;
        let offset = u64::from_le_bytes(entry[..8].try_into().unwrap());
        let length = u64::from_le_bytes(entry[8..].try_into().unwrap());
        let mut bytes = vec![0; length as usize];
        self.data.seek(SeekFrom::Start(offset)).map_err(io_error)?;
        self.data.read_exact(&mut bytes).map_err(io_error)?;
        serde_json::from_slice(&bytes).map_err(xml_error)
    }
}
fn io_error(e: impl std::fmt::Display) -> KernelError {
    KernelError::new("NATIVE_TEMPORARY_IO", e.to_string())
}
fn xml_error(e: impl std::fmt::Display) -> KernelError {
    KernelError::new("XML_INVALID", e.to_string())
}
