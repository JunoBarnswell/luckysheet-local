//! Streaming worksheet boundary. Only a single cell and one canonical row are
//! buffered. Unowned XML is copied from the original consumed bytes.
use crate::document::{CellRecord, ResourceLimits};
use kernel_core::{
    Cell, CellAddress, CellReader, KernelError, KernelResult, MAX_COLUMNS, MAX_ROWS, RangeRef,
    Scalar,
};
use quick_xml::{
    Reader,
    events::{BytesStart, Event},
};
use serde_json::{Map, Value, json};
use std::{
    collections::BTreeMap,
    io::{BufRead, Read, Write},
};

#[derive(Debug, Clone, PartialEq, serde::Serialize, serde::Deserialize)]
pub(crate) struct SharedString {
    pub text: String,
    pub rich_text: Option<Value>,
}

fn error(code: &str, message: impl Into<String>) -> KernelError {
    KernelError::new(code, message)
        .recover("Correct the source OOXML or use a supported worksheet feature")
}
fn xml_error(e: impl std::fmt::Display) -> KernelError {
    error("XML_INVALID", e.to_string())
}
fn io_error(e: impl std::fmt::Display) -> KernelError {
    error("NATIVE_IO_ERROR", e.to_string())
}
fn local(name: &[u8]) -> &[u8] {
    name.rsplit(|b| *b == b':').next().unwrap_or(name)
}
fn namespace_prefix(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|b| *b == b':')
        .map(|at| &name[..=at])
        .unwrap_or(b"")
}
fn attr(e: &BytesStart<'_>, name: &[u8]) -> KernelResult<Option<String>> {
    let mut result = None;
    for a in e.attributes() {
        let a = a.map_err(xml_error)?;
        if a.key.as_ref() == name {
            result = Some(a.unescape_value().map_err(xml_error)?.into_owned());
        }
    }
    Ok(result)
}

pub(crate) fn parse_ref(reference: &str) -> KernelResult<(u32, u32)> {
    let split = reference
        .bytes()
        .position(|b| !b.is_ascii_uppercase())
        .unwrap_or(reference.len());
    if split == 0
        || split == reference.len()
        || reference.as_bytes()[split] == b'0'
        || !reference.as_bytes()[split..].iter().all(u8::is_ascii_digit)
    {
        return Err(error(
            "OOXML_CELL_ADDRESS_INVALID",
            format!("Invalid cell reference {reference}"),
        ));
    }
    let mut column = 0u32;
    for b in reference.bytes().take(split) {
        column = column
            .checked_mul(26)
            .and_then(|n| n.checked_add((b - b'A' + 1) as u32))
            .ok_or_else(|| error("OOXML_CELL_ADDRESS_INVALID", reference))?;
    }
    let row: u32 = reference[split..]
        .parse()
        .map_err(|_| error("OOXML_CELL_ADDRESS_INVALID", reference))?;
    if row == 0 || row > MAX_ROWS || column == 0 || column > MAX_COLUMNS {
        return Err(error("OOXML_CELL_ADDRESS_INVALID", reference));
    }
    Ok((row - 1, column - 1))
}
fn reference(row: u32, column: u32) -> String {
    let mut n = column + 1;
    let mut letters = Vec::new();
    while n > 0 {
        letters.push((b'A' + ((n - 1) % 26) as u8) as char);
        n = (n - 1) / 26;
    }
    format!(
        "{}{}",
        letters.into_iter().rev().collect::<String>(),
        row + 1
    )
}

struct Recording<R> {
    inner: R,
    consumed: Vec<u8>,
}
impl<R: BufRead> Read for Recording<R> {
    fn read(&mut self, out: &mut [u8]) -> std::io::Result<usize> {
        let n = self.inner.read(out)?;
        self.consumed.extend_from_slice(&out[..n]);
        Ok(n)
    }
}
impl<R: BufRead> BufRead for Recording<R> {
    fn fill_buf(&mut self) -> std::io::Result<&[u8]> {
        self.inner.fill_buf()
    }
    fn consume(&mut self, n: usize) {
        // quick_xml consumes only bytes returned by fill_buf.
        if let Ok(buffer) = self.inner.fill_buf() {
            self.consumed.extend_from_slice(&buffer[..n]);
        }
        self.inner.consume(n);
    }
}
struct XmlStream<R: BufRead> {
    reader: Reader<Recording<R>>,
    buffer: Vec<u8>,
    pending_markup: bool,
}
impl<R: BufRead> XmlStream<R> {
    fn new(input: R) -> Self {
        Self {
            reader: Reader::from_reader(Recording {
                inner: input,
                consumed: Vec::new(),
            }),
            buffer: Vec::new(),
            pending_markup: false,
        }
    }
    fn next(&mut self) -> KernelResult<(Event<'static>, Vec<u8>)> {
        self.buffer.clear();
        let event = self
            .reader
            .read_event_into(&mut self.buffer)
            .map_err(xml_error)?
            .into_owned();
        let mut raw = std::mem::take(&mut self.reader.get_mut().consumed);
        // quick_xml's text event consumes the following opening '<'. Transfer
        // it to the following markup event, so each retained span is exact.
        if self.pending_markup {
            raw.insert(0, b'<');
            self.pending_markup = false;
        }
        if matches!(event, Event::Text(_)) && raw.last() == Some(&b'<') {
            raw.pop();
            self.pending_markup = true;
        }
        if matches!(event, Event::DocType(_)) {
            return Err(error(
                "UNSUPPORTED_FEATURE",
                "DTD declarations are not supported",
            ));
        }
        Ok((event, raw))
    }
    fn subtree(&mut self, initial: Vec<u8>, empty: bool, limit: u64) -> KernelResult<Vec<u8>> {
        let mut raw = initial;
        let mut depth = if empty { 0 } else { 1 };
        while depth > 0 {
            let (event, bytes) = self.next()?;
            match event {
                Event::Start(_) => depth += 1,
                Event::End(_) => depth -= 1,
                Event::Eof => return Err(error("XML_INVALID", "Truncated XML subtree")),
                _ => {}
            }
            if (raw.len() as u64).saturating_add(bytes.len() as u64) > limit {
                return Err(error("RESOURCE_LIMIT", "Cell XML exceeds entry budget"));
            }
            raw.extend(bytes);
        }
        Ok(raw)
    }
}

/// Read one si/is subtree, excluding phonetic guide text from the cell value.
pub(crate) fn parse_shared_string(bytes: &[u8]) -> KernelResult<SharedString> {
    let mut reader = Reader::from_reader(bytes);
    let mut buffer = Vec::new();
    let mut path: Vec<Vec<u8>> = Vec::new();
    let mut text = String::new();
    let mut runs = Vec::new();
    let mut run_text = String::new();
    let mut style = Map::new();
    let mut preserved: Vec<String> = Vec::new();
    let mut in_run = false;
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(e) | Event::Empty(e) => {
                let is_empty =
                    bytes.get(reader.buffer_position().saturating_sub(2) as usize) == Some(&b'/');
                let name = local(e.name().as_ref()).to_vec();
                if path.len() == 1 && name == b"r" {
                    in_run = true;
                    run_text.clear();
                    style.clear();
                    preserved.clear();
                }
                if path.last().is_some_and(|p| p == b"rPr") {
                    parse_run_property(&e, &name, &mut style, &mut preserved)?;
                }
                if !is_empty {
                    path.push(name);
                }
            }
            Event::Text(e) => {
                if path.last().is_some_and(|p| p == b"t") && !path.iter().any(|p| p == b"rPh") {
                    let value = e.unescape().map_err(xml_error)?;
                    text.push_str(&value);
                    if in_run {
                        run_text.push_str(&value);
                    }
                }
            }
            Event::CData(e) => {
                if path.last().is_some_and(|p| p == b"t") && !path.iter().any(|p| p == b"rPh") {
                    let value = e.decode().map_err(xml_error)?;
                    text.push_str(&value);
                    if in_run {
                        run_text.push_str(&value);
                    }
                }
            }
            Event::End(e) => {
                if local(e.name().as_ref()) == b"r" && path.len() == 2 {
                    let mut run = json!({"text":run_text});
                    if !style.is_empty() {
                        run["style"] = Value::Object(style.clone());
                    }
                    if !preserved.is_empty() {
                        run["preservedProperties"] = json!(preserved);
                    }
                    runs.push(run);
                    in_run = false;
                }
                path.pop();
            }
            Event::DocType(_) => return Err(error("UNSUPPORTED_FEATURE", "DTD in string")),
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !path.is_empty() {
        return Err(error("XML_INVALID", "Truncated rich string"));
    }
    if !runs.is_empty()
        && runs
            .iter()
            .filter_map(|r| r["text"].as_str())
            .collect::<String>()
            != text
    {
        return Err(error(
            "OOXML_STRING_INVALID",
            "Mixed plain and rich string content",
        ));
    }
    Ok(SharedString {
        text,
        rich_text: if runs.is_empty() {
            None
        } else {
            Some(Value::Array(runs))
        },
    })
}
fn parse_run_property(
    e: &BytesStart<'_>,
    name: &[u8],
    style: &mut Map<String, Value>,
    preserved: &mut Vec<String>,
) -> KernelResult<()> {
    let value = attr(e, b"val")?;
    match name {
        b"b" | b"i" | b"strike" | b"u" => {
            let key = match name {
                b"b" => "bold",
                b"i" => "italic",
                b"strike" => "strikethrough",
                _ => "underline",
            };
            style.insert(
                key.into(),
                json!(!matches!(
                    value.as_deref(),
                    Some("0" | "false" | "off" | "none")
                )),
            );
            if name == b"u" && !matches!(value.as_deref(), None | Some("single" | "none")) {
                preserved.push("u".into());
            }
        }
        b"rFont" => {
            style.insert(
                "fontFamily".into(),
                json!(value.ok_or_else(|| error("OOXML_STRING_INVALID", "Missing rFont val"))?),
            );
        }
        b"sz" => {
            let points: f64 = value
                .ok_or_else(|| error("OOXML_STRING_INVALID", "Missing font size"))?
                .parse()
                .map_err(xml_error)?;
            if !points.is_finite() || points <= 0.0 {
                return Err(error("OOXML_STRING_INVALID", "Invalid font size"));
            }
            style.insert("fontSizePx".into(), json!(points * 96.0 / 72.0));
        }
        b"color" => {
            if let Some(rgb) = attr(e, b"rgb")? {
                if !matches!(rgb.len(), 6 | 8) || !rgb.bytes().all(|b| b.is_ascii_hexdigit()) {
                    return Err(error("OOXML_STRING_INVALID", "Invalid RGB color"));
                }
                style.insert(
                    "textColor".into(),
                    json!(format!("#{}", &rgb[rgb.len() - 6..])),
                );
            } else {
                preserved.push("color".into());
            }
        }
        b"vertAlign" => {
            let value =
                value.ok_or_else(|| error("OOXML_STRING_INVALID", "Missing vertAlign val"))?;
            if !["baseline", "superscript", "subscript"].contains(&value.as_str()) {
                return Err(error("OOXML_STRING_INVALID", "Invalid vertAlign"));
            }
            style.insert("verticalAlignment".into(), json!(value));
        }
        _ => preserved.push(String::from_utf8_lossy(name).into_owned()),
    }
    Ok(())
}

#[derive(Default)]
struct FormulaGroups {
    shared: BTreeMap<u32, (u32, u32, String, String)>,
}

fn parse_cell(
    sheet: &str,
    bytes: &[u8],
    shared: &mut dyn FnMut(usize) -> KernelResult<SharedString>,
    styles: &[Value],
    groups: &mut FormulaGroups,
) -> KernelResult<CellRecord> {
    let mut stream = XmlStream::new(bytes);
    let (event, _) = stream.next()?;
    let (start, empty) = match event {
        Event::Start(e) => (e, false),
        Event::Empty(e) => (e, true),
        _ => return Err(error("XML_INVALID", "Expected cell")),
    };
    let reference = attr(&start, b"r")?
        .ok_or_else(|| error("OOXML_CELL_ADDRESS_MISSING", "Cell r is required"))?;
    let (row, column) = parse_ref(&reference)?;
    let kind = attr(&start, b"t")?.unwrap_or_else(|| "n".into());
    let mut cell = Cell::default();
    if let Some(index) = attr(&start, b"s")? {
        let index: usize = index.parse().map_err(xml_error)?;
        let style = styles.get(index).ok_or_else(|| {
            error(
                "OOXML_STYLE_INDEX_INVALID",
                format!("Style {index} does not exist"),
            )
        })?;
        cell.metadata
            .insert("styleId".into(), json!(format!("ooxml:{index}")));
        if style.as_object().is_none_or(|s| !s.is_empty()) {
            cell.metadata.insert("style".into(), style.clone());
        }
    }
    let mut value = None;
    let mut inline = None;
    let mut formula = None;
    if !empty {
        loop {
            let (event, raw) = stream.next()?;
            match event {
                Event::Start(e) | Event::Empty(e) => {
                    let is_empty = raw.ends_with(b"/>");
                    let name = local(e.name().as_ref()).to_vec();
                    let owned = namespace_prefix(e.name().as_ref())
                        == namespace_prefix(start.name().as_ref());
                    let node = stream.subtree(raw, is_empty, u64::MAX)?;
                    if !owned {
                        continue;
                    }
                    match name.as_slice() {
                        b"v" => {
                            if value.is_some() {
                                return Err(error("OOXML_CELL_INVALID", "Duplicate value"));
                            }
                            value = Some(element_text(&node)?);
                        }
                        b"is" => {
                            if inline.is_some() {
                                return Err(error("OOXML_CELL_INVALID", "Duplicate inline string"));
                            }
                            inline = Some(parse_shared_string(&node)?);
                        }
                        b"f" => {
                            if formula.is_some() {
                                return Err(error("OOXML_CELL_INVALID", "Duplicate formula"));
                            }
                            let family = attr(&e, b"t")?.unwrap_or_else(|| "normal".into());
                            let text = element_text(&node)?;
                            match family.as_str() {
                                "normal" => {
                                    if text.trim().is_empty() {
                                        return Err(error(
                                            "OOXML_FORMULA_INVALID",
                                            "Empty normal formula",
                                        ));
                                    }
                                    formula = Some(format!("={text}"));
                                }
                                "shared" => {
                                    let index: u32 = attr(&e, b"si")?
                                        .ok_or_else(|| {
                                            error("OOXML_FORMULA_INVALID", "Shared si missing")
                                        })?
                                        .parse()
                                        .map_err(xml_error)?;
                                    let master = !text.is_empty();
                                    if master {
                                        let range = attr(&e, b"ref")?.ok_or_else(|| {
                                            error(
                                                "OOXML_FORMULA_INVALID",
                                                "Shared master range missing",
                                            )
                                        })?;
                                        if groups.shared.contains_key(&index) {
                                            return Err(error(
                                                "OOXML_FORMULA_INVALID",
                                                "Duplicate shared master",
                                            ));
                                        }
                                        validate_group_range(&range, row, column)?;
                                        groups.shared.insert(
                                            index,
                                            (row, column, format!("={text}"), range),
                                        );
                                    }
                                    let (master_row, master_column, source, range) = groups.shared.get(&index).ok_or_else(|| error("OOXML_FORMULA_INVALID", format!("Shared master si={index} must precede dependent cell")))?;
                                    validate_group_range(range, row, column)?;
                                    formula = Some(kernel_formula::references::offset(
                                        source,
                                        row as i32 - *master_row as i32,
                                        column as i32 - *master_column as i32,
                                    )?);
                                    cell.metadata.insert("formulaMetadata".into(), json!({"kind":"shared","sharedIndex":index,"sharedMaster":master,"range":range}));
                                }
                                "array" => {
                                    if text.is_empty() {
                                        return Err(error(
                                            "OOXML_FORMULA_INVALID",
                                            "Array formula is empty",
                                        ));
                                    }
                                    let range = attr(&e, b"ref")?.ok_or_else(|| {
                                        error("OOXML_FORMULA_INVALID", "Array range missing")
                                    })?;
                                    validate_group_range(&range, row, column)?;
                                    formula = Some(format!("={text}"));
                                    cell.metadata.insert(
                                        "formulaMetadata".into(),
                                        json!({"kind":"array","range":range}),
                                    );
                                }
                                _ => {
                                    return Err(error(
                                        "UNSUPPORTED_FEATURE",
                                        format!(
                                            "{sheet}!{reference}: {family} formula is unsupported"
                                        ),
                                    ));
                                }
                            }
                        }
                        _ => {}
                    }
                }
                Event::End(_) => break,
                Event::Eof => return Err(error("XML_INVALID", "Truncated cell")),
                _ => {}
            }
        }
    }
    cell.value = match kind.as_str() {
        "inlineStr" => {
            if value.is_some() || formula.is_some() {
                return Err(error(
                    "OOXML_CELL_INVALID",
                    "inlineStr must contain only is",
                ));
            }
            let string =
                inline.ok_or_else(|| error("OOXML_STRING_INVALID", "inlineStr is missing is"))?;
            if let Some(rich) = string.rich_text {
                cell.metadata.insert("richText".into(), rich);
            }
            Scalar::Text(string.text)
        }
        "s" => {
            if inline.is_some() {
                return Err(error(
                    "OOXML_CELL_INVALID",
                    "Shared cell contains inline string",
                ));
            }
            let index: usize = value
                .ok_or_else(|| error("OOXML_STRING_INVALID", "Shared string index missing"))?
                .parse()
                .map_err(xml_error)?;
            let string = shared(index)?;
            if let Some(rich) = string.rich_text {
                cell.metadata.insert("richText".into(), rich);
            }
            Scalar::Text(string.text)
        }
        "n" | "b" | "str" | "e" => {
            if inline.is_some() {
                return Err(error("OOXML_CELL_INVALID", "Unexpected inline string"));
            }
            match value {
                None => Scalar::Null,
                Some(v) if kind == "str" => Scalar::Text(v),
                Some(v) if kind == "e" => Scalar::error(&v, "Native formula error"),
                Some(v) if kind == "b" => match v.as_str() {
                    "1" | "true" => Scalar::Boolean(true),
                    "0" | "false" => Scalar::Boolean(false),
                    _ => return Err(error("OOXML_VALUE_INVALID", "Invalid Boolean")),
                },
                Some(v) if v.is_empty() => Scalar::Null,
                Some(v) => Scalar::Number(
                    v.parse()
                        .map_err(|_| error("OOXML_VALUE_INVALID", format!("Invalid number {v}")))?,
                ),
            }
        }
        _ => {
            return Err(error(
                "UNSUPPORTED_FEATURE",
                format!("Cell type {kind} is unsupported"),
            ));
        }
    };
    cell.value.validate()?;
    cell.formula = formula;
    Ok(CellRecord {
        address: CellAddress {
            sheet_id: sheet.into(),
            row,
            column,
        },
        cell,
    })
}
fn element_text(bytes: &[u8]) -> KernelResult<String> {
    let mut reader = Reader::from_reader(bytes);
    let mut buffer = Vec::new();
    let mut result = String::new();
    let mut depth = 0;
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(_) => {
                depth += 1;
                if depth > 1 {
                    return Err(error(
                        "OOXML_CELL_INVALID",
                        "Nested content in scalar/formula",
                    ));
                }
            }
            Event::End(_) => {
                depth -= 1;
            }
            Event::Text(e) => result.push_str(&e.unescape().map_err(xml_error)?),
            Event::CData(e) => result.push_str(&e.decode().map_err(xml_error)?),
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    Ok(result)
}

pub(crate) fn parse_sheet<R: BufRead>(
    sheet_id: &str,
    input: R,
    shared: &mut dyn FnMut(usize) -> KernelResult<SharedString>,
    styles: &[Value],
    limits: &ResourceLimits,
    visitor: &mut dyn FnMut(CellRecord) -> KernelResult<()>,
) -> KernelResult<u64> {
    let mut stream = XmlStream::new(input);
    let mut path = Vec::<Vec<u8>>::new();
    let mut count = 0u64;
    let mut row = None;
    let mut previous_row = None;
    let mut previous_column = None;
    let mut sheet_data_seen = false;
    let mut groups = FormulaGroups::default();
    let mut prefix = Vec::new();
    loop {
        let (event, raw) = stream.next()?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = raw.ends_with(b"/>");
                let name = local(e.name().as_ref()).to_vec();
                if path.is_empty() {
                    if name != b"worksheet" {
                        return Err(error("OOXML_SHEET_INVALID", "Expected worksheet root"));
                    }
                    prefix = namespace_prefix(e.name().as_ref()).to_vec();
                } else if namespace_prefix(e.name().as_ref()) != prefix {
                    stream.subtree(raw, empty, limits.max_entry_bytes)?;
                    continue;
                }
                if path.len() == 1 && name == b"sheetData" {
                    if sheet_data_seen {
                        return Err(error("OOXML_SHEET_INVALID", "Duplicate sheetData"));
                    }
                    sheet_data_seen = true;
                }
                if path.len() == 2 && path[1] == b"sheetData" && name == b"row" {
                    let current = row_number(&e)?;
                    if previous_row.is_some_and(|r| r >= current) {
                        return Err(error(
                            "OOXML_ROW_ORDER_INVALID",
                            "Rows must have unique increasing references",
                        ));
                    }
                    row = Some(current);
                    previous_row = Some(current);
                    previous_column = None;
                }
                if path.len() == 3 && path[1] == b"sheetData" && path[2] == b"row" && name == b"c" {
                    let raw = stream.subtree(raw, empty, limits.max_entry_bytes)?;
                    let record = parse_cell(sheet_id, &raw, shared, styles, &mut groups)?;
                    if row != Some(record.address.row)
                        || previous_column.is_some_and(|c| c >= record.address.column)
                    {
                        return Err(error(
                            "OOXML_CELL_ORDER_INVALID",
                            "Cell reference must match its row and increase within the row",
                        ));
                    }
                    previous_column = Some(record.address.column);
                    count += 1;
                    if count > limits.max_cells {
                        return Err(error(
                            "CELL_LIMIT",
                            "Worksheet cell count exceeds the budget",
                        ));
                    }
                    visitor(record)?;
                    continue;
                }
                if !empty {
                    path.push(name);
                }
            }
            Event::End(_) => {
                path.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !path.is_empty() || !sheet_data_seen {
        return Err(error(
            "OOXML_SHEET_INVALID",
            "Missing or truncated sheetData",
        ));
    }
    Ok(count)
}
fn validate_group_range(range: &str, row: u32, column: u32) -> KernelResult<()> {
    let (first, last) = range.split_once(':').unwrap_or((range, range));
    let (sr, sc) = parse_ref(first)?;
    let (er, ec) = parse_ref(last)?;
    if sr > er || sc > ec || row < sr || row > er || column < sc || column > ec {
        return Err(error(
            "OOXML_FORMULA_INVALID",
            "Formula group reference does not contain the cell",
        ));
    }
    Ok(())
}
fn row_number(e: &BytesStart<'_>) -> KernelResult<u32> {
    let raw = attr(e, b"r")?.ok_or_else(|| error("OOXML_ROW_INVALID", "Row r is required"))?;
    let n: u32 = raw.parse().map_err(|_| error("OOXML_ROW_INVALID", &raw))?;
    if n == 0 || n > MAX_ROWS {
        return Err(error("OOXML_ROW_INVALID", raw));
    }
    Ok(n - 1)
}

/// Canonical reads are bounded to one row; their callback order is irrelevant.
/// row_count is the canonical manifest extent, not a source dimension hint.
pub(crate) fn rewrite_sheet<R: BufRead, W: Write>(
    sheet_id: &str,
    input: R,
    mut output: W,
    cells: &dyn CellReader,
    row_count: u32,
    shared: &mut dyn FnMut(usize) -> KernelResult<SharedString>,
    source_styles: &[Value],
    style_index: &mut dyn FnMut(&Cell) -> KernelResult<Option<u32>>,
) -> KernelResult<()> {
    if row_count > MAX_ROWS {
        return Err(error(
            "RANGE_INVALID",
            "Canonical row count exceeds Excel bounds",
        ));
    }
    let dimension = canonical_dimension(sheet_id, cells, row_count)?;
    let mut stream = XmlStream::new(input);
    let mut path = Vec::<Vec<u8>>::new();
    let mut row_cells = BTreeMap::new();
    let mut current_row = None;
    let mut next_row = 0;
    let mut data_seen = false;
    let mut prefix = String::new();
    let mut groups = FormulaGroups::default();
    let mut arrays: Vec<String> = Vec::new();
    loop {
        let (event, raw) = stream.next()?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let empty = raw.ends_with(b"/>");
                let name = local(e.name().as_ref()).to_vec();
                if path.is_empty() && name == b"worksheet" {
                    if let Some(at) = e.name().as_ref().iter().position(|b| *b == b':') {
                        prefix = String::from_utf8(e.name().as_ref()[..=at].to_vec())
                            .map_err(xml_error)?;
                    }
                }
                if !path.is_empty() && namespace_prefix(e.name().as_ref()) != prefix.as_bytes() {
                    let bytes = stream.subtree(raw, empty, u64::MAX)?;
                    write_bytes(&mut output, &bytes)?;
                    continue;
                }
                if path.len() == 1 && name == b"dimension" {
                    let raw = stream.subtree(raw, empty, u64::MAX)?;
                    write_open(&mut output, &e, &[("ref", Some(dimension.clone()))])?;
                    if empty {
                        write_bytes(&mut output, format!("</{prefix}dimension>").as_bytes())?;
                    } else {
                        let mut retained = XmlStream::new(raw.as_slice());
                        retained.next()?;
                        loop {
                            let (event, bytes) = retained.next()?;
                            if matches!(event, Event::Eof) {
                                break;
                            }
                            write_bytes(&mut output, &bytes)?;
                        }
                    }
                    continue;
                }
                if path.len() == 1 && name == b"sheetData" {
                    if data_seen {
                        return Err(error("OOXML_SHEET_INVALID", "Duplicate sheetData"));
                    }
                    data_seen = true;
                    if empty {
                        write_open(&mut output, &e, &[])?;
                        emit_rows(
                            sheet_id,
                            cells,
                            0,
                            row_count,
                            &prefix,
                            &mut output,
                            style_index,
                            &arrays,
                        )?;
                        write_bytes(&mut output, format!("</{}sheetData>", prefix).as_bytes())?;
                        next_row = row_count;
                        continue;
                    }
                }
                if path.len() == 2 && path[1] == b"sheetData" && name == b"row" {
                    let row = row_number(&e)?;
                    if row < next_row {
                        return Err(error(
                            "OOXML_ROW_ORDER_INVALID",
                            "Rows are not strictly ordered",
                        ));
                    }
                    emit_rows(
                        sheet_id,
                        cells,
                        next_row,
                        row.min(row_count),
                        &prefix,
                        &mut output,
                        style_index,
                        &arrays,
                    )?;
                    next_row = row + 1;
                    current_row = Some(row);
                    row_cells = read_row(sheet_id, cells, row)?;
                    if empty && !row_cells.is_empty() {
                        reject_array_insertions(row, &row_cells, &arrays)?;
                        write_open(&mut output, &e, &[])?;
                        emit_cells(
                            row,
                            std::mem::take(&mut row_cells),
                            &prefix,
                            &mut output,
                            style_index,
                        )?;
                        write_bytes(&mut output, format!("</{}row>", prefix).as_bytes())?;
                        current_row = None;
                        continue;
                    }
                }
                if path.len() == 3 && path[1] == b"sheetData" && path[2] == b"row" && name == b"c" {
                    let raw = stream.subtree(raw, empty, u64::MAX)?;
                    let source = parse_cell(sheet_id, &raw, shared, source_styles, &mut groups)?;
                    if current_row != Some(source.address.row) {
                        return Err(error("OOXML_CELL_ORDER_INVALID", "Cell outside owning row"));
                    }
                    if source
                        .cell
                        .metadata
                        .get("formulaMetadata")
                        .and_then(|v| v["kind"].as_str())
                        == Some("array")
                    {
                        arrays.push(
                            source.cell.metadata["formulaMetadata"]["range"]
                                .as_str()
                                .ok_or_else(|| {
                                    error("OOXML_FORMULA_INVALID", "Array range missing")
                                })?
                                .into(),
                        );
                    }
                    let after = row_cells.split_off(&source.address.column);
                    reject_array_insertions(source.address.row, &row_cells, &arrays)?;
                    emit_cells(
                        source.address.row,
                        std::mem::replace(&mut row_cells, after),
                        &prefix,
                        &mut output,
                        style_index,
                    )?;
                    let canonical = row_cells.remove(&source.address.column);
                    if canonical.as_ref().is_some_and(|c| c == &source.cell) {
                        write_bytes(&mut output, &raw)?;
                    } else {
                        let changed_content = canonical.as_ref().is_none_or(|c| {
                            c.value != source.cell.value
                                || c.formula != source.cell.formula
                                || c.metadata.get("richText")
                                    != source.cell.metadata.get("richText")
                        });
                        if changed_content
                            && (arrays.iter().any(|range| {
                                validate_group_range(
                                    range,
                                    source.address.row,
                                    source.address.column,
                                )
                                .is_ok()
                            }) || source
                                .cell
                                .metadata
                                .get("formulaMetadata")
                                .and_then(|v| v["kind"].as_str())
                                == Some("shared"))
                        {
                            return Err(error(
                                "UNSUPPORTED_FEATURE",
                                format!(
                                    "{}!{}: formula group edits require atomic group replacement",
                                    sheet_id,
                                    reference(source.address.row, source.address.column)
                                ),
                            ));
                        }
                        rewrite_cell(
                            &raw,
                            &source,
                            canonical.as_ref(),
                            &prefix,
                            &mut output,
                            style_index,
                        )?;
                    }
                    continue;
                }
                write_bytes(&mut output, &raw)?;
                if !empty {
                    path.push(name);
                }
            }
            Event::End(_) => {
                if path.len() == 3 && path[1] == b"sheetData" && path[2] == b"row" {
                    reject_array_insertions(
                        current_row
                            .ok_or_else(|| error("OOXML_ROW_INVALID", "Row owner missing"))?,
                        &row_cells,
                        &arrays,
                    )?;
                    emit_cells(
                        current_row
                            .ok_or_else(|| error("OOXML_ROW_INVALID", "Row owner missing"))?,
                        std::mem::take(&mut row_cells),
                        &prefix,
                        &mut output,
                        style_index,
                    )?;
                    current_row = None;
                }
                if path.len() == 2 && path[1] == b"sheetData" {
                    emit_rows(
                        sheet_id,
                        cells,
                        next_row,
                        row_count,
                        &prefix,
                        &mut output,
                        style_index,
                        &arrays,
                    )?;
                }
                write_bytes(&mut output, &raw)?;
                path.pop();
            }
            Event::Eof => break,
            _ => write_bytes(&mut output, &raw)?,
        }
    }
    if !data_seen || !path.is_empty() {
        return Err(error(
            "OOXML_SHEET_INVALID",
            "Missing or truncated sheetData",
        ));
    }
    Ok(())
}
fn canonical_dimension(
    sheet: &str,
    cells: &dyn CellReader,
    row_count: u32,
) -> KernelResult<String> {
    if row_count == 0 {
        return Ok("A1".into());
    }
    let mut bounds: Option<(u32, u32, u32, u32)> = None;
    cells.read_range(
        &RangeRef {
            sheet_id: sheet.into(),
            start_row: 0,
            end_row: row_count - 1,
            start_column: 0,
            end_column: MAX_COLUMNS - 1,
        },
        &mut |address, _| {
            if address.sheet_id != sheet
                || address.row >= row_count
                || address.column >= MAX_COLUMNS
            {
                return Err(error(
                    "CELL_READER_CONTRACT_INVALID",
                    "Dimension reader returned a cell outside its range",
                ));
            }
            bounds = Some(match bounds {
                Some((sr, sc, er, ec)) => (
                    sr.min(address.row),
                    sc.min(address.column),
                    er.max(address.row),
                    ec.max(address.column),
                ),
                None => (address.row, address.column, address.row, address.column),
            });
            Ok(())
        },
    )?;
    Ok(match bounds {
        Some((sr, sc, er, ec)) if sr != er || sc != ec => {
            format!("{}:{}", reference(sr, sc), reference(er, ec))
        }
        Some((sr, sc, _, _)) => reference(sr, sc),
        None => "A1".into(),
    })
}
fn read_row(sheet: &str, cells: &dyn CellReader, row: u32) -> KernelResult<BTreeMap<u32, Cell>> {
    let mut result = BTreeMap::new();
    cells.read_range(
        &RangeRef {
            sheet_id: sheet.into(),
            start_row: row,
            end_row: row,
            start_column: 0,
            end_column: MAX_COLUMNS - 1,
        },
        &mut |address, cell| {
            if address.sheet_id != sheet
                || address.row != row
                || address.column >= MAX_COLUMNS
                || result.insert(address.column, cell).is_some()
            {
                return Err(error(
                    "CELL_READER_CONTRACT_INVALID",
                    "Row reader returned duplicate or out-of-range cell",
                ));
            }
            Ok(())
        },
    )?;
    Ok(result)
}
fn emit_rows<W: Write>(
    sheet: &str,
    cells: &dyn CellReader,
    start: u32,
    end: u32,
    prefix: &str,
    output: &mut W,
    style_index: &mut dyn FnMut(&Cell) -> KernelResult<Option<u32>>,
    arrays: &[String],
) -> KernelResult<()> {
    for row in start..end {
        let values = read_row(sheet, cells, row)?;
        if values.is_empty() {
            continue;
        }
        reject_array_insertions(row, &values, arrays)?;
        write_bytes(
            output,
            format!("<{prefix}row r=\"{}\">", row + 1).as_bytes(),
        )?;
        for (column, cell) in values {
            write_new_cell(row, column, &cell, prefix, output, style_index)?;
        }
        write_bytes(output, format!("</{prefix}row>").as_bytes())?;
    }
    Ok(())
}
fn reject_array_insertions(
    row: u32,
    cells: &BTreeMap<u32, Cell>,
    arrays: &[String],
) -> KernelResult<()> {
    if cells.keys().any(|column| {
        arrays
            .iter()
            .any(|range| validate_group_range(range, row, *column).is_ok())
    }) {
        return Err(error(
            "UNSUPPORTED_FEATURE",
            "Inserting cells inside an array requires atomic group replacement",
        ));
    }
    Ok(())
}
// Address is carried explicitly by the row emitter, including inserted cells.
fn emit_cells<W: Write>(
    row: u32,
    values: BTreeMap<u32, Cell>,
    prefix: &str,
    output: &mut W,
    style_index: &mut dyn FnMut(&Cell) -> KernelResult<Option<u32>>,
) -> KernelResult<()> {
    for (column, cell) in values {
        write_new_cell(row, column, &cell, prefix, output, style_index)?;
    }
    Ok(())
}
fn write_bytes<W: Write>(output: &mut W, bytes: &[u8]) -> KernelResult<()> {
    if bytes
        .iter()
        .any(|b| *b < 0x20 && ![b'\t', b'\n', b'\r'].contains(b))
    {
        return Err(error(
            "XML_INVALID",
            "XML contains a forbidden control character",
        ));
    }
    output.write_all(bytes).map_err(io_error)
}
fn escaped(text: &str) -> String {
    quick_xml::escape::escape(text).into_owned()
}
fn write_open<W: Write>(
    output: &mut W,
    start: &BytesStart<'_>,
    replace: &[(&str, Option<String>)],
) -> KernelResult<()> {
    write_bytes(output, b"<")?;
    write_bytes(output, start.name().as_ref())?;
    // Validate attributes, but copy every unowned attribute lexeme exactly,
    // including its quote choice, entity spelling, and leading whitespace.
    for a in start.attributes() {
        a.map_err(xml_error)?;
    }
    let raw = start.as_ref();
    let mut at = start.name().as_ref().len();
    while at < raw.len() {
        let begin = at;
        while at < raw.len() && raw[at].is_ascii_whitespace() {
            at += 1;
        }
        if at == raw.len() {
            write_bytes(output, &raw[begin..])?;
            break;
        }
        let key_start = at;
        while at < raw.len() && raw[at] != b'=' && !raw[at].is_ascii_whitespace() {
            at += 1;
        }
        let key = &raw[key_start..at];
        while at < raw.len() && raw[at] != b'=' {
            at += 1;
        }
        at += 1;
        while at < raw.len() && raw[at].is_ascii_whitespace() {
            at += 1;
        }
        let quote = *raw
            .get(at)
            .ok_or_else(|| error("XML_INVALID", "Attribute quote missing"))?;
        at += 1;
        while at < raw.len() && raw[at] != quote {
            at += 1;
        }
        if at >= raw.len() {
            return Err(error("XML_INVALID", "Unclosed attribute"));
        }
        at += 1;
        if !replace.iter().any(|(owned, _)| owned.as_bytes() == key) {
            write_bytes(output, &raw[begin..at])?;
        }
    }
    for (key, value) in replace {
        if let Some(value) = value {
            write_bytes(output, format!(" {key}=\"{}\"", escaped(value)).as_bytes())?;
        }
    }
    write_bytes(output, b">")
}
fn cell_type(cell: &Cell) -> Option<&'static str> {
    match &cell.value {
        Scalar::Text(_) => Some(if cell.formula.is_some() {
            "str"
        } else {
            "inlineStr"
        }),
        Scalar::Boolean(_) => Some("b"),
        Scalar::Error(_) => Some("e"),
        _ => None,
    }
}
fn write_new_cell<W: Write>(
    row: u32,
    column: u32,
    cell: &Cell,
    prefix: &str,
    output: &mut W,
    style_index: &mut dyn FnMut(&Cell) -> KernelResult<Option<u32>>,
) -> KernelResult<()> {
    let name = format!("{prefix}c");
    let start = BytesStart::new(&name);
    let index = style_index(cell)?.map(|i| i.to_string());
    write_open(
        output,
        &start,
        &[
            ("r", Some(reference(row, column))),
            ("t", cell_type(cell).map(str::to_owned)),
            ("s", index),
        ],
    )?;
    write_content(cell, prefix, output)?;
    write_bytes(output, format!("</{name}>").as_bytes())
}
fn rewrite_cell<W: Write>(
    raw: &[u8],
    source: &CellRecord,
    canonical: Option<&Cell>,
    prefix: &str,
    output: &mut W,
    style_index: &mut dyn FnMut(&Cell) -> KernelResult<Option<u32>>,
) -> KernelResult<()> {
    let mut stream = XmlStream::new(raw);
    let (event, _) = stream.next()?;
    let (start, empty) = match event {
        Event::Start(e) => (e, false),
        Event::Empty(e) => (e, true),
        _ => return Err(error("XML_INVALID", "Expected cell")),
    };
    let replaces_string = canonical.is_none_or(|c| {
        c.value != source.cell.value
            || c.formula != source.cell.formula
            || c.metadata.get("richText") != source.cell.metadata.get("richText")
    });
    if replaces_string {
        validate_inline_replacement(raw)?;
    }
    if canonical.is_none() {
        for a in start.attributes() {
            let a = a.map_err(xml_error)?;
            if ![b"r".as_slice(), b"s", b"t"].contains(&a.key.as_ref()) {
                return Err(error(
                    "UNSUPPORTED_FEATURE",
                    "Deleting a cell with preserved attributes requires explicit removal ownership",
                ));
            }
        }
        if !empty {
            loop {
                let (event, bytes) = stream.next()?;
                match event {
                    Event::Start(e) | Event::Empty(e) => {
                        let is_empty = bytes.ends_with(b"/>");
                        if namespace_prefix(e.name().as_ref())
                            != namespace_prefix(start.name().as_ref())
                            || ![b"f".as_slice(), b"v", b"is"].contains(&local(e.name().as_ref()))
                        {
                            return Err(error(
                                "UNSUPPORTED_FEATURE",
                                "Deleting a cell with preserved XML requires explicit removal ownership",
                            ));
                        }
                        stream.subtree(bytes, is_empty, u64::MAX)?;
                    }
                    Event::End(_) => break,
                    Event::Text(e) if e.unescape().map_err(xml_error)?.trim().is_empty() => {}
                    _ => {
                        return Err(error(
                            "UNSUPPORTED_FEATURE",
                            "Deleting a cell with preserved XML requires explicit removal ownership",
                        ));
                    }
                }
            }
        }
        return Ok(());
    }
    let same_content = canonical.is_some_and(|c| {
        c.value == source.cell.value
            && c.formula == source.cell.formula
            && c.metadata.get("richText") == source.cell.metadata.get("richText")
            && c.metadata.get("formulaMetadata") == source.cell.metadata.get("formulaMetadata")
    });
    if !same_content
        && source
            .cell
            .metadata
            .get("richText")
            .and_then(Value::as_array)
            .is_some_and(|runs| {
                runs.iter().any(|r| {
                    r.get("preservedProperties")
                        .and_then(Value::as_array)
                        .is_some_and(|p| !p.is_empty())
                })
            })
    {
        return Err(error(
            "UNSUPPORTED_FEATURE",
            "Replacing rich text with preserved source properties requires explicit removal ownership",
        ));
    }
    let index = canonical
        .map(&mut *style_index)
        .transpose()?
        .flatten()
        .map(|i| i.to_string());
    let mut replacements = vec![("s", index)];
    if !same_content {
        replacements.push(("t", canonical.and_then(cell_type).map(str::to_owned)));
    }
    write_open(output, &start, &replacements)?;
    if !same_content {
        if let Some(cell) = canonical {
            write_content(cell, prefix, output)?;
        }
    }
    if !empty {
        loop {
            let (event, bytes) = stream.next()?;
            match event {
                Event::Start(e) | Event::Empty(e) => {
                    let is_empty = bytes.ends_with(b"/>");
                    let owned = namespace_prefix(e.name().as_ref())
                        == namespace_prefix(start.name().as_ref())
                        && [b"f".as_slice(), b"v", b"is"].contains(&local(e.name().as_ref()));
                    let bytes = stream.subtree(bytes, is_empty, u64::MAX)?;
                    if same_content || !owned {
                        write_bytes(output, &bytes)?;
                    }
                }
                Event::End(_) => break,
                Event::Eof => return Err(error("XML_INVALID", "Truncated cell")),
                _ => write_bytes(output, &bytes)?,
            }
        }
    }
    write_bytes(
        output,
        format!("</{}>", String::from_utf8_lossy(start.name().as_ref())).as_bytes(),
    )
}
fn validate_inline_replacement(raw: &[u8]) -> KernelResult<()> {
    let mut stream = XmlStream::new(raw);
    let mut path: Vec<Vec<u8>> = Vec::new();
    let mut prefix = Vec::new();
    loop {
        let (event, bytes) = stream.next()?;
        match event {
            Event::Start(e) | Event::Empty(e) => {
                let name = local(e.name().as_ref()).to_vec();
                if path.is_empty() {
                    prefix = namespace_prefix(e.name().as_ref()).to_vec();
                }
                if path.iter().any(|n| n == b"is") || name == b"is" {
                    let allowed_attributes: &[&[u8]] = match name.as_slice() {
                        b"is" | b"r" | b"rPr" => &[],
                        b"t" => &[b"xml:space"],
                        b"b" | b"i" | b"strike" | b"u" | b"rFont" | b"sz" | b"vertAlign" => {
                            &[b"val"]
                        }
                        b"color" => &[b"rgb"],
                        _ => {
                            return Err(error(
                                "UNSUPPORTED_FEATURE",
                                "Inline text contains preserved phonetic or extension nodes",
                            ));
                        }
                    };
                    if namespace_prefix(e.name().as_ref()) != prefix {
                        return Err(error(
                            "UNSUPPORTED_FEATURE",
                            "Inline text contains preserved namespaces",
                        ));
                    }
                    for a in e.attributes() {
                        let a = a.map_err(xml_error)?;
                        if !allowed_attributes.contains(&a.key.as_ref()) {
                            return Err(error(
                                "UNSUPPORTED_FEATURE",
                                "Inline text contains preserved attributes",
                            ));
                        }
                    }
                }
                if !bytes.ends_with(b"/>") {
                    path.push(name);
                }
            }
            Event::End(_) => {
                path.pop();
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(())
}
fn write_content<W: Write>(cell: &Cell, prefix: &str, output: &mut W) -> KernelResult<()> {
    cell.value.validate()?;
    if cell
        .metadata
        .get("formulaMetadata")
        .and_then(|v| v["kind"].as_str())
        .is_some_and(|kind| kind != "normal")
    {
        return Err(error(
            "UNSUPPORTED_FEATURE",
            "Creating formula groups requires an atomic group contract",
        ));
    }
    if let Some(formula) = &cell.formula {
        let formula = formula
            .strip_prefix('=')
            .ok_or_else(|| error("FORMULA_INVALID", "Canonical formula must begin with ="))?;
        if formula.is_empty() {
            return Err(error("FORMULA_INVALID", "Formula is empty"));
        }
        write_bytes(
            output,
            format!("<{prefix}f>{}</{prefix}f>", escaped(formula)).as_bytes(),
        )?;
    }
    let scalar = match &cell.value {
        Scalar::Null => return Ok(()),
        Scalar::Boolean(v) => {
            if *v {
                "1".into()
            } else {
                "0".into()
            }
        }
        Scalar::Number(v) => v.to_string(),
        Scalar::Error(v) => v.code.clone(),
        Scalar::Text(text) if cell.formula.is_none() => {
            write_bytes(output, format!("<{prefix}is>").as_bytes())?;
            if let Some(runs) = cell.metadata.get("richText") {
                let runs = runs
                    .as_array()
                    .ok_or_else(|| error("CELL_VALUE_INVALID", "richText must be an array"))?;
                if runs
                    .iter()
                    .filter_map(|r| r["text"].as_str())
                    .collect::<String>()
                    != *text
                {
                    return Err(error(
                        "CELL_VALUE_INVALID",
                        "richText plain projection differs from value",
                    ));
                }
                for run in runs {
                    if run
                        .get("preservedProperties")
                        .and_then(Value::as_array)
                        .is_some_and(|v| !v.is_empty())
                    {
                        return Err(error(
                            "UNSUPPORTED_FEATURE",
                            "Editing preserved rich text properties is unsupported",
                        ));
                    }
                    write_bytes(output, format!("<{prefix}r>").as_bytes())?;
                    if let Some(style) = run.get("style") {
                        write_run_style(style, prefix, output)?;
                    }
                    let text = run["text"]
                        .as_str()
                        .ok_or_else(|| error("CELL_VALUE_INVALID", "Rich run text missing"))?;
                    write_bytes(
                        output,
                        format!(
                            "<{prefix}t xml:space=\"preserve\">{}</{prefix}t></{prefix}r>",
                            escaped(text)
                        )
                        .as_bytes(),
                    )?;
                }
            } else {
                write_bytes(
                    output,
                    format!(
                        "<{prefix}t xml:space=\"preserve\">{}</{prefix}t>",
                        escaped(text)
                    )
                    .as_bytes(),
                )?;
            }
            return write_bytes(output, format!("</{prefix}is>").as_bytes());
        }
        Scalar::Text(v) => v.clone(),
    };
    write_bytes(
        output,
        format!("<{prefix}v>{}</{prefix}v>", escaped(&scalar)).as_bytes(),
    )
}
fn write_run_style<W: Write>(style: &Value, prefix: &str, output: &mut W) -> KernelResult<()> {
    let style = style
        .as_object()
        .ok_or_else(|| error("CELL_VALUE_INVALID", "Rich style must be an object"))?;
    write_bytes(output, format!("<{prefix}rPr>").as_bytes())?;
    for (key, value) in style {
        let (name, attribute, text) = match key.as_str() {
            "bold" | "italic" | "underline" | "strikethrough" => {
                let enabled = value
                    .as_bool()
                    .ok_or_else(|| error("CELL_VALUE_INVALID", "Rich Boolean style invalid"))?;
                let tag = match key.as_str() {
                    "bold" => "b",
                    "italic" => "i",
                    "underline" => "u",
                    _ => "strike",
                };
                (
                    tag,
                    "val",
                    if key == "underline" {
                        if enabled { "single" } else { "none" }
                    } else if enabled {
                        "1"
                    } else {
                        "0"
                    }
                    .to_owned(),
                )
            }
            "fontFamily" => (
                "rFont",
                "val",
                value
                    .as_str()
                    .ok_or_else(|| error("CELL_VALUE_INVALID", "Font family invalid"))?
                    .into(),
            ),
            "fontSizePx" => (
                "sz",
                "val",
                (value
                    .as_f64()
                    .filter(|n| n.is_finite() && *n > 0.0)
                    .ok_or_else(|| error("CELL_VALUE_INVALID", "Font size invalid"))?
                    * 72.0
                    / 96.0)
                    .to_string(),
            ),
            "textColor" => {
                let color = value
                    .as_str()
                    .and_then(|s| s.strip_prefix('#'))
                    .filter(|s| s.len() == 6 && s.bytes().all(|b| b.is_ascii_hexdigit()))
                    .ok_or_else(|| error("UNSUPPORTED_FEATURE", "Rich font color must be RGB"))?;
                ("color", "rgb", format!("FF{color}"))
            }
            "verticalAlignment" => (
                "vertAlign",
                "val",
                value
                    .as_str()
                    .filter(|s| ["baseline", "superscript", "subscript"].contains(s))
                    .ok_or_else(|| error("CELL_VALUE_INVALID", "Rich vertical alignment invalid"))?
                    .into(),
            ),
            _ => {
                return Err(error(
                    "UNSUPPORTED_FEATURE",
                    format!("Rich style {key} unsupported"),
                ));
            }
        };
        write_bytes(
            output,
            format!("<{prefix}{name} {attribute}=\"{}\"/>", escaped(&text)).as_bytes(),
        )?;
    }
    write_bytes(output, format!("</{prefix}rPr>").as_bytes())
}

#[cfg(test)]
mod tests {
    use super::*;
    fn no_strings(_: usize) -> KernelResult<SharedString> {
        Err(error("OOXML_STRING_INDEX_INVALID", "No shared strings"))
    }
    fn parse(bytes: &[u8]) -> KernelResult<Vec<CellRecord>> {
        let mut values = Vec::new();
        parse_sheet(
            "s",
            bytes,
            &mut no_strings,
            &[],
            &ResourceLimits::default(),
            &mut |cell| {
                values.push(cell);
                Ok(())
            },
        )?;
        Ok(values)
    }
    struct Cells(Vec<CellRecord>);
    impl CellReader for Cells {
        fn revision(&self) -> u64 {
            1
        }
        fn read_cell(&self, address: &CellAddress) -> KernelResult<Option<Cell>> {
            Ok(self
                .0
                .iter()
                .find(|c| c.address == *address)
                .map(|c| c.cell.clone()))
        }
        fn read_range(
            &self,
            range: &RangeRef,
            visitor: &mut dyn FnMut(CellAddress, Cell) -> KernelResult<()>,
        ) -> KernelResult<()> {
            for cell in self.0.iter().rev().filter(|c| range.contains(&c.address)) {
                visitor(cell.address.clone(), cell.cell.clone())?;
            }
            Ok(())
        }
    }
    #[test]
    fn parses_entities_rich_strings_empty_cells_and_shared_formula() {
        let xml = br#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/><sz val="12"/></rPr><t xml:space="preserve"> a&amp;b </t></r><r><t>&lt;x&gt;</t></r><rPh sb="0" eb="1"><t>guide</t></rPh></is></c><c r="B1"><f t="shared" si="0" ref="B1:C1">A1+$A$2</f><v>3</v></c><c r="C1"><f t="shared" si="0"/><v>4</v></c><c r="D1"/></row></sheetData></worksheet>"#;
        let cells = parse(xml).unwrap();
        assert_eq!(cells[0].cell.value, Scalar::Text(" a&b <x>".into()));
        assert_eq!(
            cells[0].cell.metadata["richText"][0]["style"]["fontSizePx"],
            16.0
        );
        assert_eq!(cells[2].cell.formula.as_deref(), Some("=B1+$A$2"));
        assert_eq!(cells[3].cell.value, Scalar::Null);
    }
    #[test]
    fn rejects_invalid_references_types_rows_and_numbers() {
        for cell in [
            "<c r=\"A0\"/>",
            "<c r=\"XFE1\"/>",
            "<c r=\"A2\"/>",
            "<c r=\"A1\" t=\"b\"><v>7</v></c>",
            "<c r=\"A1\"><v>NaN</v></c>",
            "<c r=\"A1\" t=\"d\"><v>2026-01-01</v></c>",
        ] {
            assert!(
                parse(
                    format!(
                        "<worksheet><sheetData><row r=\"1\">{cell}</row></sheetData></worksheet>"
                    )
                    .as_bytes()
                )
                .is_err(),
                "{cell}"
            );
        }
        assert!(parse(b"<worksheet><sheetData><row r=\"0\"/></sheetData></worksheet>").is_err());
    }
    #[test]
    fn rewrites_owned_content_and_inserts_cells_rows_preserving_unknown_bytes() {
        let xml = br#"<?xml version='1.0'?><worksheet xmlns:z='urn:unknown'><z:root a='&#38;'>  untouched </z:root><sheetData><row r='1' custom='&#65;'><c r='A1' z:flag = '&#x41;'><v>1</v><z:node foo='&#38;'> preserved </z:node></c><c r='C1'><v>3</v></c></row></sheetData><z:tail/></worksheet>"#;
        let mut cells = parse(xml).unwrap();
        cells[0].cell.value = Scalar::Number(2.0);
        cells.push(CellRecord {
            address: CellAddress {
                sheet_id: "s".into(),
                row: 0,
                column: 1,
            },
            cell: Cell {
                value: Scalar::Text("&".into()),
                ..Cell::default()
            },
        });
        cells.push(CellRecord {
            address: CellAddress {
                sheet_id: "s".into(),
                row: 2,
                column: 0,
            },
            cell: Cell {
                value: Scalar::Boolean(true),
                ..Cell::default()
            },
        });
        let mut output = Vec::new();
        rewrite_sheet(
            "s",
            xml.as_slice(),
            &mut output,
            &Cells(cells),
            3,
            &mut no_strings,
            &[],
            &mut |_| Ok(None),
        )
        .unwrap();
        let result = String::from_utf8(output.clone()).unwrap();
        assert!(result.contains("z:flag = '&#x41;'"));
        assert!(result.contains("<z:root a='&#38;'>  untouched </z:root>"));
        assert!(result.contains("<z:node foo='&#38;'> preserved </z:node>"));
        let cells = parse(&output).unwrap();
        assert_eq!(cells.len(), 4);
        assert_eq!(cells[0].cell.value, Scalar::Number(2.0));
        assert_eq!(cells[1].cell.value, Scalar::Text("&".into()));
        assert_eq!(cells[3].address.row, 2);
    }
    #[test]
    fn unmodified_cells_are_byte_identical_and_deleted_cells_do_not_reappear() {
        let xml = br#"<worksheet><sheetData><row r='1'><c r='A1' t='inlineStr'><is><r><rPr><u val='double'/></rPr><t>x</t></r></is></c><c r='B1'><v>2</v></c></row></sheetData></worksheet>"#;
        let mut cells = parse(xml).unwrap();
        cells.pop();
        let mut output = Vec::new();
        rewrite_sheet(
            "s",
            xml.as_slice(),
            &mut output,
            &Cells(cells),
            1,
            &mut no_strings,
            &[],
            &mut |_| Ok(None),
        )
        .unwrap();
        assert!(String::from_utf8_lossy(&output).contains(
            "<c r='A1' t='inlineStr'><is><r><rPr><u val='double'/></rPr><t>x</t></r></is></c>"
        ));
        assert_eq!(parse(&output).unwrap().len(), 1);
    }
    #[test]
    fn formula_groups_reject_partial_edits() {
        let xml = br#"<worksheet><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:B1">1+1</f><v>2</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#;
        let mut cells = parse(xml).unwrap();
        cells[1].cell.value = Scalar::Number(3.0);
        assert!(
            rewrite_sheet(
                "s",
                xml.as_slice(),
                Vec::new(),
                &Cells(cells),
                1,
                &mut no_strings,
                &[],
                &mut |_| Ok(None)
            )
            .is_err()
        );
    }
    #[test]
    fn whitespace_before_cells_and_unknown_cell_namespaces_remain_exact() {
        let xml = b"<worksheet xmlns:z='urn:other'>\n<sheetData>\n <row r='1'>\n  <c r='A1'><v>1</v><z:v>unowned</z:v></c>\n </row>\n</sheetData>\n</worksheet>";
        let mut cells = parse(xml).unwrap();
        cells[0].cell.value = Scalar::Number(5.0);
        let mut output = Vec::new();
        rewrite_sheet(
            "s",
            xml.as_slice(),
            &mut output,
            &Cells(cells),
            1,
            &mut no_strings,
            &[],
            &mut |_| Ok(None),
        )
        .unwrap();
        assert!(String::from_utf8_lossy(&output).contains("<z:v>unowned</z:v>"));
        assert_eq!(parse(&output).unwrap()[0].cell.value, Scalar::Number(5.0));
    }
    #[test]
    fn editing_inline_phonetic_content_fails_without_dropping_source_xml() {
        let xml = br#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t><rPh sb="0" eb="1"><t>guide</t></rPh></is></c></row></sheetData></worksheet>"#;
        let mut cells = parse(xml).unwrap();
        cells[0].cell.value = Scalar::Text("y".into());
        assert!(
            rewrite_sheet(
                "s",
                xml.as_slice(),
                Vec::new(),
                &Cells(cells),
                1,
                &mut no_strings,
                &[],
                &mut |_| Ok(None)
            )
            .is_err()
        );
    }
}
