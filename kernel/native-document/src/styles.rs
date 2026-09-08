use crate::xmlnode::{self, error, escape, Node};
use kernel_core::{Cell, KernelResult};
use serde_json::{json, Map, Value};
use std::collections::BTreeMap;

/// Native stylesheet owner. Existing records and unknown children retain their
/// exact bytes; authored formats append real font/fill/border/xf records.
#[derive(Debug, Clone)]
pub(crate) struct Styles {
    source: Vec<u8>,
    root: Node,
    pub values: Vec<Value>,
    additions: BTreeMap<String, Vec<String>>,
    num_formats: BTreeMap<String, u32>,
    next_num_format: u32,
    unsupported: Vec<bool>,
}
impl Styles {
    pub fn parse(source: Vec<u8>, theme: Option<&[u8]>, budget: u64) -> KernelResult<Self> {
        let root = xmlnode::parse(&source, budget)?;
        if root.local() != "styleSheet" {
            return Err(error(
                "OOXML_STYLES_INVALID",
                "Styles relationship does not target a stylesheet",
            ));
        }
        let colors = theme
            .map(|bytes| theme_colors(bytes, budget))
            .transpose()?
            .unwrap_or_default();
        let font_records = root
            .child("fonts")
            .map(|n| {
                n.children_named("font")
                    .map(|n| Ok((font(n, &colors)?, font_has_unmodeled_content(n))))
                    .collect::<KernelResult<Vec<_>>>()
            })
            .transpose()?
            .unwrap_or_default();
        let (fonts, font_unsupported): (Vec<_>, Vec<_>) = font_records.into_iter().unzip();
        let fill_records = root
            .child("fills")
            .map(|n| {
                n.children_named("fill")
                    .map(|n| Ok((fill(n, &colors)?, fill_has_unmodeled_content(n))))
                    .collect::<KernelResult<Vec<_>>>()
            })
            .transpose()?
            .unwrap_or_default();
        let (fills, fill_unsupported): (Vec<_>, Vec<_>) = fill_records.into_iter().unzip();
        let border_records = root
            .child("borders")
            .map(|n| {
                n.children_named("border")
                    .map(|n| Ok((border(n, &colors)?, border_has_unmodeled_content(n))))
                    .collect::<KernelResult<Vec<_>>>()
            })
            .transpose()?
            .unwrap_or_default();
        let (borders, border_unsupported): (Vec<_>, Vec<_>) =
            border_records.into_iter().unzip();
        let mut formats = builtin_formats();
        let mut next_num_format = 164;
        if let Some(parent) = root.child("numFmts") {
            for n in parent.children_named("numFmt") {
                let id = uint(n, "numFmtId", 0)?;
                let code = n
                    .attr("formatCode")
                    .ok_or_else(|| error("OOXML_STYLES_INVALID", "Number format has no code"))?
                    .to_owned();
                formats.insert(id, code);
                next_num_format = next_num_format.max(id + 1);
            }
        }
        let num_formats = formats
            .iter()
            .map(|(id, code)| (code.clone(), *id))
            .collect();
        let mut values = Vec::new();
        let mut unsupported = Vec::new();
        if let Some(parent) = root.child("cellXfs") {
            for xf in parent.children_named("xf") {
                let mut style = Map::new();
                let mut style_unsupported = has_unmodeled_attributes(
                    xf,
                    &[
                        "numFmtId",
                        "fontId",
                        "fillId",
                        "borderId",
                        "xfId",
                        "applyNumberFormat",
                        "applyFont",
                        "applyFill",
                        "applyBorder",
                        "applyAlignment",
                        "applyProtection",
                    ],
                );
                for (field, items) in [
                    ("fontId", &fonts),
                    ("fillId", &fills),
                    ("borderId", &borders),
                ] {
                    let index = uint(xf, field, 0)? as usize;
                    if let Some(item) = items.get(index) {
                        style.extend(item.clone());
                        style_unsupported |= match field {
                            "fontId" => font_unsupported.get(index).copied().unwrap_or(false),
                            "fillId" => fill_unsupported.get(index).copied().unwrap_or(false),
                            "borderId" => border_unsupported.get(index).copied().unwrap_or(false),
                            _ => false,
                        };
                    } else if index != 0 || !items.is_empty() {
                        return Err(error(
                            "OOXML_STYLE_INDEX_INVALID",
                            format!("{field} {index}"),
                        ));
                    }
                }
                let fmt = uint(xf, "numFmtId", 0)?;
                if let Some(code) = formats.get(&fmt) {
                    if fmt != 0 {
                        style.insert("numberFormat".into(), json!(code));
                    }
                } else {
                    return Err(error("UNSUPPORTED_FEATURE",format!("Locale-dependent built-in number format {fmt} has no explicit format code")));
                }
                if let Some(alignment) = xf.child("alignment") {
                    alignment_style(alignment, &mut style)?;
                    style_unsupported |= alignment_has_unmodeled_content(alignment);
                }
                if let Some(protection) = xf.child("protection") {
                    style.insert("locked".into(), json!(boolean(protection, "locked", true)?));
                    style.insert(
                        "formulaHidden".into(),
                        json!(boolean(protection, "hidden", false)?),
                    );
                    style_unsupported |= has_unmodeled_attributes(protection, &["locked", "hidden"])
                        || !protection.children.is_empty();
                }
                style_unsupported |= xf
                    .children
                    .iter()
                    .any(|c| !["alignment", "protection"].contains(&c.local()));
                unsupported.push(style_unsupported);
                values.push(Value::Object(style));
            }
        }
        if values.is_empty() {
            return Err(error(
                "OOXML_STYLES_INVALID",
                "Stylesheet must declare cellXfs",
            ));
        }
        Ok(Self {
            source,
            root,
            values,
            additions: BTreeMap::new(),
            num_formats,
            next_num_format,
            unsupported,
        })
    }
    pub fn default_styles() -> KernelResult<Self> {
        Self::parse(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#.to_vec(),None,16*1024*1024)
    }
    pub fn index_for(&mut self, cell: &Cell) -> KernelResult<Option<u32>> {
        let Some(value) = cell.metadata.get("style") else {
            return Ok(None);
        };
        let style = value
            .as_object()
            .ok_or_else(|| error("CELL_STYLE_INVALID", "style must be an object"))?;
        if let Some(index) = cell
            .metadata
            .get("styleId")
            .and_then(Value::as_str)
            .and_then(|s| s.strip_prefix("ooxml:"))
            .and_then(|s| s.parse::<usize>().ok())
        {
            if self.values.get(index) == Some(value) {
                return Ok(Some(index as u32));
            }
            if self.unsupported.get(index) == Some(&true) {
                return Err(error(
                    "UNSUPPORTED_FEATURE",
                    "Changing a style with unknown native extensions is not owned",
                ));
            }
        }
        if let Some(index) = self.values.iter().position(|candidate| candidate == value) {
            return Ok(Some(index as u32));
        }
        let font_id = self.add("fonts", "font", write_font(style)?)?;
        let fill_id = self.add("fills", "fill", write_fill(style)?)?;
        let border_id = self.add("borders", "border", write_border(style)?)?;
        let mut fmt = 0;
        if let Some(code) = style.get("numberFormat").and_then(Value::as_str) {
            if let Some(id) = self.num_formats.get(code) {
                fmt = *id;
            } else {
                fmt = self.next_num_format;
                self.next_num_format += 1;
                self.add(
                    "numFmts",
                    "numFmt",
                    format!(
                        "<numFmt numFmtId=\"{fmt}\" formatCode=\"{}\"/>",
                        escape(code)
                    ),
                )?;
                self.num_formats.insert(code.into(), fmt);
            }
        }
        let alignment = write_alignment(style)?;
        let protection = if style.contains_key("locked") || style.contains_key("formulaHidden") {
            format!(
                "<protection locked=\"{}\" hidden=\"{}\"/>",
                flag(style, "locked", true)?,
                flag(style, "formulaHidden", false)?
            )
        } else {
            String::new()
        };
        let xf=format!("<xf numFmtId=\"{fmt}\" fontId=\"{font_id}\" fillId=\"{fill_id}\" borderId=\"{border_id}\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyNumberFormat=\"1\" applyAlignment=\"1\" applyProtection=\"1\">{alignment}{protection}</xf>");
        let index = self.add("cellXfs", "xf", xf)?;
        self.values.push(value.clone());
        self.unsupported.push(false);
        Ok(Some(index))
    }
    fn add(&mut self, collection: &str, child: &str, value: String) -> KernelResult<u32> {
        let count = self
            .root
            .child(collection)
            .map(|n| n.children_named(child).count())
            .unwrap_or(0);
        let list = self.additions.entry(collection.into()).or_default();
        let index = count + list.len();
        if index >= 65_490 {
            return Err(error("STYLE_LIMIT", "Stylesheet record budget exceeded"));
        }
        list.push(value);
        Ok(index as u32)
    }
    pub fn changed(&self) -> bool {
        !self.additions.is_empty()
    }
    pub fn bytes(&self) -> KernelResult<Vec<u8>> {
        if !self.changed() {
            return Ok(self.source.clone());
        }
        let mut edits: Vec<(usize, usize, Vec<u8>)> = Vec::new();
        for (name, additions) in &self.additions {
            if let Some(node) = self.root.child(name) {
                let tag = match name.as_str() {
                    "fonts" => "font",
                    "fills" => "fill",
                    "borders" => "border",
                    "cellXfs" => "xf",
                    "numFmts" => "numFmt",
                    _ => return Err(error("STYLES_INTERNAL", "Unowned collection")),
                };
                let count = node.children_named(tag).count() + additions.len();
                let mut opening = format!("<{}", node.name);
                for (key, value) in &node.attributes {
                    if key != "count" {
                        opening.push_str(&format!(" {key}=\"{}\"", escape(value)));
                    }
                }
                opening.push_str(&format!(" count=\"{count}\">"));
                let mut replacement = opening.into_bytes();
                if !node.empty {
                    replacement
                        .extend_from_slice(&self.source[node.opening_end..node.closing_start]);
                }
                for addition in additions {
                    replacement.extend_from_slice(addition.as_bytes());
                }
                replacement.extend_from_slice(format!("</{}>", node.name).as_bytes());
                edits.push((node.span.start, node.span.end, replacement));
            } else {
                let insertion = self
                    .root
                    .children
                    .first()
                    .map(|n| n.span.start)
                    .unwrap_or(self.root.closing_start);
                edits.push((
                    insertion,
                    insertion,
                    format!(
                        "<{name} count=\"{}\">{}</{name}>",
                        additions.len(),
                        additions.join("")
                    )
                    .into_bytes(),
                ));
            }
        }
        edits.sort_by_key(|e| e.0);
        let mut out = Vec::new();
        let mut position = 0;
        for (start, end, replacement) in edits {
            out.extend_from_slice(&self.source[position..start]);
            out.extend_from_slice(&replacement);
            position = end;
        }
        out.extend_from_slice(&self.source[position..]);
        Ok(out)
    }
}
fn uint(node: &Node, key: &str, default: u32) -> KernelResult<u32> {
    node.attr(key)
        .map(|v| {
            v.parse()
                .map_err(|_| error("OOXML_ATTRIBUTE_INVALID", format!("{}@{key}", node.name)))
        })
        .unwrap_or(Ok(default))
}
fn boolean(node: &Node, key: &str, default: bool) -> KernelResult<bool> {
    match node.attr(key) {
        None => Ok(default),
        Some("1" | "true") => Ok(true),
        Some("0" | "false") => Ok(false),
        _ => Err(error(
            "OOXML_ATTRIBUTE_INVALID",
            format!("{}@{key}", node.name),
        )),
    }
}
fn numeric(node: &Node, key: &str) -> KernelResult<Option<f64>> {
    node.attr(key)
        .map(|v| {
            v.parse::<f64>()
                .map_err(|_| error("OOXML_ATTRIBUTE_INVALID", key))
                .and_then(|n| {
                    if n.is_finite() {
                        Ok(n)
                    } else {
                        Err(error("OOXML_ATTRIBUTE_INVALID", key))
                    }
                })
        })
        .transpose()
}
fn theme_colors(bytes: &[u8], budget: u64) -> KernelResult<Vec<String>> {
    let root = xmlnode::parse(bytes, budget)?;
    let scheme = root
        .child("themeElements")
        .and_then(|n| n.child("clrScheme"))
        .ok_or_else(|| error("OOXML_THEME_INVALID", "Theme has no color scheme"))?;
    let mut values = Vec::new();
    for key in [
        "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5",
        "accent6", "hlink", "folHlink",
    ] {
        let node = scheme
            .child(key)
            .and_then(|n| n.children.first())
            .ok_or_else(|| error("OOXML_THEME_INVALID", key))?;
        values.push(
            node.attr("lastClr")
                .or(node.attr("val"))
                .ok_or_else(|| error("OOXML_THEME_INVALID", key))?
                .into(),
        );
    }
    Ok(values)
}
fn color(node: &Node, theme: &[String]) -> KernelResult<Option<String>> {
    let raw = if let Some(rgb) = node.attr("rgb") {
        Some(rgb.to_owned())
    } else if let Some(id) = node.attr("theme") {
        let id = id
            .parse::<usize>()
            .map_err(|_| error("OOXML_COLOR_INVALID", id))?;
        Some(
            theme
                .get(id)
                .ok_or_else(|| {
                    error(
                        "OOXML_THEME_MISSING",
                        format!("Theme color {id} has no declared theme"),
                    )
                })?
                .clone(),
        )
    } else if let Some(id) = node.attr("indexed") {
        let id = id
            .parse::<usize>()
            .map_err(|_| error("OOXML_COLOR_INVALID", id))?;
        INDEXED.get(id).map(|s| s.to_string())
    } else {
        None
    };
    let Some(raw) = raw else { return Ok(None) };
    let raw = raw.trim_start_matches('#');
    let raw = if raw.len() == 8 { &raw[2..] } else { raw };
    if raw.len() != 6 || !raw.bytes().all(|c| c.is_ascii_hexdigit()) {
        return Err(error("OOXML_COLOR_INVALID", raw));
    }
    let tint = numeric(node, "tint")?.unwrap_or(0.0);
    if !(-1.0..=1.0).contains(&tint) {
        return Err(error("OOXML_COLOR_INVALID", "Tint outside [-1,1]"));
    }
    let mut channels = [0.0; 3];
    for (i, c) in channels.iter_mut().enumerate() {
        *c = u8::from_str_radix(&raw[i * 2..i * 2 + 2], 16)
            .map_err(|_| error("OOXML_COLOR_INVALID", raw))? as f64
            / 255.0;
    }
    // OOXML tint transforms HLS luminance, not independent RGB channels.
    let max = channels.iter().copied().fold(0.0, f64::max);
    let min = channels.iter().copied().fold(1.0, f64::min);
    let luminance = (max + min) / 2.0;
    let target = if tint < 0.0 {
        luminance * (1.0 + tint)
    } else {
        luminance * (1.0 - tint) + tint
    };
    if tint != 0.0 {
        let delta = max - min;
        let saturation = if delta == 0.0 {
            0.0
        } else {
            delta / (1.0 - (2.0 * luminance - 1.0).abs())
        };
        let hue = if delta == 0.0 {
            0.0
        } else if max == channels[0] {
            ((channels[1] - channels[2]) / delta).rem_euclid(6.0)
        } else if max == channels[1] {
            (channels[2] - channels[0]) / delta + 2.0
        } else {
            (channels[0] - channels[1]) / delta + 4.0
        };
        let chroma = (1.0 - (2.0 * target - 1.0).abs()) * saturation;
        let x = chroma * (1.0 - (hue.rem_euclid(2.0) - 1.0).abs());
        let m = target - chroma / 2.0;
        channels = match hue as u32 {
            0 => [chroma, x, 0.0],
            1 => [x, chroma, 0.0],
            2 => [0.0, chroma, x],
            3 => [0.0, x, chroma],
            4 => [x, 0.0, chroma],
            _ => [chroma, 0.0, x],
        };
        for c in &mut channels {
            *c += m;
        }
    }
    Ok(Some(format!(
        "#{:02X}{:02X}{:02X}",
        (channels[0] * 255.0).round() as u8,
        (channels[1] * 255.0).round() as u8,
        (channels[2] * 255.0).round() as u8
    )))
}

fn has_unmodeled_attributes(node: &Node, allowed: &[&str]) -> bool {
    node.attributes
        .keys()
        .any(|key| !allowed.contains(&key.as_str()))
}

fn has_unmodeled_children(node: &Node, allowed: &[&str]) -> bool {
    node.children
        .iter()
        .any(|child| !allowed.contains(&child.local()))
}

fn color_has_unmodeled_content(node: &Node) -> bool {
    has_unmodeled_attributes(node, &["rgb", "theme", "indexed", "tint"])
        || !node.children.is_empty()
}

fn font_has_unmodeled_content(node: &Node) -> bool {
    if has_unmodeled_attributes(node, &[]) {
        return true;
    }
    node.children.iter().any(|child| match child.local() {
        "name" | "sz" | "b" | "i" | "strike" | "u" | "scheme" => {
            has_unmodeled_attributes(child, &["val"]) || !child.children.is_empty()
        }
        "vertAlign" => {
            has_unmodeled_attributes(child, &["val"])
                || !child.children.is_empty()
                || !matches!(child.attr("val"), Some("superscript" | "subscript"))
        }
        "color" => color_has_unmodeled_content(child),
        _ => true,
    })
}

fn fill_has_unmodeled_content(node: &Node) -> bool {
    if has_unmodeled_attributes(node, &[]) || node.children.len() > 1 {
        return true;
    }
    let Some(fill) = node.children.first() else {
        return false;
    };
    match fill.local() {
        "patternFill" => {
            if has_unmodeled_attributes(fill, &["patternType"])
                || has_unmodeled_children(fill, &["fgColor", "bgColor"])
            {
                return true;
            }
            let pattern = fill.attr("patternType").unwrap_or("none");
            ((pattern == "none" || pattern == "gray125") && !fill.children.is_empty())
                || fill
                    .children
                    .iter()
                    .any(color_has_unmodeled_content)
        }
        "gradientFill" => {
            if has_unmodeled_attributes(fill, &["type", "degree"])
            {
                return true;
            }
            fill.children.iter().any(|stop| {
                stop.local() != "stop"
                    || has_unmodeled_attributes(stop, &["position"])
                    || stop.children.len() != 1
                    || stop.children
                        .first()
                        .is_some_and(color_has_unmodeled_content)
                        || stop.children.first().is_none_or(|child| child.local() != "color")
            })
        }
        _ => true,
    }
}

fn border_has_unmodeled_content(node: &Node) -> bool {
    if has_unmodeled_attributes(node, &["diagonalUp", "diagonalDown"])
        || has_unmodeled_children(node, &["top", "bottom", "left", "right", "diagonal"])
    {
        return true;
    }
    node.children.iter().any(|side| {
        has_unmodeled_attributes(side, &["style"])
            || side.children.iter().any(|child| {
                child.local() != "color" || color_has_unmodeled_content(child)
            })
            || (side.attr("style").is_none() && !side.children.is_empty())
    })
}

fn alignment_has_unmodeled_content(node: &Node) -> bool {
    has_unmodeled_attributes(
        node,
        &[
            "horizontal",
            "vertical",
            "wrapText",
            "shrinkToFit",
            "indent",
            "textRotation",
            "readingOrder",
        ],
    ) || !node.children.is_empty()
        || node
            .attr("indent")
            .is_some_and(|value| value.parse::<u32>().is_err())
}

const INDEXED: [&str; 64] = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000",
    "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000",
    "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC",
    "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF",
    "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF",
    "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600",
    "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399",
    "333333",
];
fn font(n: &Node, theme: &[String]) -> KernelResult<Map<String, Value>> {
    let mut m = Map::new();
    for c in &n.children {
        match c.local() {
            "name" => {
                if let Some(v) = c.attr("val") {
                    m.insert("fontFamily".into(), json!(v));
                }
            }
            "sz" => {
                if let Some(v) = numeric(c, "val")? {
                    m.insert("fontSizePx".into(), json!(v * 96.0 / 72.0));
                }
            }
            "b" | "i" | "strike" => {
                m.insert(
                    match c.local() {
                        "b" => "bold",
                        "i" => "italic",
                        _ => "strikethrough",
                    }
                    .into(),
                    json!(boolean(c, "val", true)?),
                );
            }
            "u" => {
                let v = c.attr("val").unwrap_or("single");
                m.insert("underline".into(), json!(v != "none"));
                if v != "none" {
                    m.insert("underlineStyle".into(), json!(v));
                }
            }
            "vertAlign" => {
                if let Some(v) = c.attr("val") {
                    if v == "superscript" || v == "subscript" {
                        m.insert(v.into(), json!(true));
                    }
                }
            }
            "color" => {
                if let Some(v) = color(c, theme)? {
                    m.insert("textColor".into(), json!(v));
                }
            }
            "scheme" => {
                if let Some(v) = c.attr("val") {
                    m.insert("fontTheme".into(), json!(v));
                }
            }
            _ => {}
        }
    }
    Ok(m)
}
fn fill(n: &Node, theme: &[String]) -> KernelResult<Map<String, Value>> {
    let mut m = Map::new();
    if let Some(p) = n.child("patternFill") {
        let pattern = p.attr("patternType").unwrap_or("none");
        if pattern != "none" && pattern != "gray125" {
            let fg = p
                .child("fgColor")
                .map(|n| color(n, theme))
                .transpose()?
                .flatten();
            let bg = p
                .child("bgColor")
                .map(|n| color(n, theme))
                .transpose()?
                .flatten();
            let mut f = Map::new();
            f.insert(
                "kind".into(),
                json!(if pattern == "solid" {
                    "solid"
                } else {
                    "pattern"
                }),
            );
            f.insert("pattern".into(), json!(pattern));
            if let Some(v) = fg {
                if pattern == "solid" {
                    m.insert("background".into(), json!(v));
                }
                f.insert("foreground".into(), json!(v));
            }
            if let Some(v) = bg {
                f.insert("background".into(), json!(v));
            }
            m.insert("fill".into(), Value::Object(f));
        }
    } else if let Some(g) = n.child("gradientFill") {
        let mut stops = Vec::new();
        for stop in g.children_named("stop") {
            let position = numeric(stop, "position")?
                .ok_or_else(|| error("OOXML_FILL_INVALID", "Missing gradient position"))?;
            let c = stop
                .children
                .first()
                .map(|n| color(n, theme))
                .transpose()?
                .flatten()
                .ok_or_else(|| error("OOXML_FILL_INVALID", "Missing gradient color"))?;
            stops.push(json!({"position":position,"color":c}));
        }
        m.insert("fill".into(),json!({"kind":"gradient","gradientType":g.attr("type").unwrap_or("linear"),"degree":numeric(g,"degree")?.unwrap_or(0.0),"stops":stops}));
    }
    Ok(m)
}
fn border(n: &Node, theme: &[String]) -> KernelResult<Map<String, Value>> {
    let mut m = Map::new();
    let mut b = Map::new();
    for name in ["top", "bottom", "left", "right", "diagonal"] {
        if let Some(side) = n.child(name) {
            if let Some(style) = side.attr("style") {
                let c = side
                    .child("color")
                    .map(|n| color(n, theme))
                    .transpose()?
                    .flatten()
                    .unwrap_or_else(|| "#000000".into());
                b.insert(name.into(), json!({"style":style,"color":c}));
            }
        }
    }
    for key in ["diagonalUp", "diagonalDown"] {
        if boolean(n, key, false)? {
            b.insert(key.into(), json!(true));
        }
    }
    if !b.is_empty() {
        m.insert("borders".into(), Value::Object(b));
    }
    Ok(m)
}
fn alignment_style(n: &Node, m: &mut Map<String, Value>) -> KernelResult<()> {
    for (native, key) in [
        ("horizontal", "horizontalAlignment"),
        ("vertical", "verticalAlignment"),
    ] {
        if let Some(v) = n.attr(native) {
            m.insert(key.into(), json!(v));
        }
    }
    for key in ["wrapText", "shrinkToFit"] {
        if n.attr(key).is_some() {
            m.insert(key.into(), json!(boolean(n, key, false)?));
        }
    }
    if let Some(v) = numeric(n, "indent")? {
        m.insert("indent".into(), json!(v));
    }
    if n.attr("textRotation").is_some() {
        let v = uint(n, "textRotation", 0)?;
        if v == 255 {
            m.insert("textOrientation".into(), json!("stacked"));
        } else if v <= 180 {
            m.insert(
                "textRotate".into(),
                json!(if v > 90 { 90 - (v as i32) } else { v as i32 }),
            );
        } else {
            return Err(error("OOXML_ALIGNMENT_INVALID", "Invalid text rotation"));
        }
    }
    if n.attr("readingOrder").is_some() {
        let v = uint(n, "readingOrder", 0)?;
        let direction = match v {
            0 => "context",
            1 => "ltr",
            2 => "rtl",
            _ => return Err(error("OOXML_ALIGNMENT_INVALID", "Invalid reading order")),
        };
        m.insert("readingOrder".into(), json!(direction));
    }
    Ok(())
}
fn builtin_formats() -> BTreeMap<u32, String> {
    [
        (0, "General"),
        (1, "0"),
        (2, "0.00"),
        (3, "#,##0"),
        (4, "#,##0.00"),
        (9, "0%"),
        (10, "0.00%"),
        (11, "0.00E+00"),
        (12, "# ?/?"),
        (13, "# ??/??"),
        (14, "mm-dd-yy"),
        (15, "d-mmm-yy"),
        (16, "d-mmm"),
        (17, "mmm-yy"),
        (18, "h:mm AM/PM"),
        (19, "h:mm:ss AM/PM"),
        (20, "h:mm"),
        (21, "h:mm:ss"),
        (22, "m/d/yy h:mm"),
        (37, "#,##0 ;(#,##0)"),
        (38, "#,##0 ;[Red](#,##0)"),
        (39, "#,##0.00;(#,##0.00)"),
        (40, "#,##0.00;[Red](#,##0.00)"),
        (45, "mm:ss"),
        (46, "[h]:mm:ss"),
        (47, "mmss.0"),
        (48, "##0.0E+0"),
        (49, "@"),
    ]
    .into_iter()
    .map(|(k, v)| (k, v.into()))
    .collect()
}
fn flag(m: &Map<String, Value>, key: &str, default: bool) -> KernelResult<u8> {
    m.get(key)
        .map(|v| {
            v.as_bool()
                .map(u8::from)
                .ok_or_else(|| error("CELL_STYLE_INVALID", key))
        })
        .unwrap_or(Ok(u8::from(default)))
}
fn rgb(value: &str) -> KernelResult<String> {
    let v = value.trim_start_matches('#');
    if v.len() != 6 || !v.bytes().all(|b| b.is_ascii_hexdigit()) {
        return Err(error(
            "CELL_STYLE_INVALID",
            format!("Color {value} must be canonical RGB"),
        ));
    }
    Ok(format!("FF{}", v.to_ascii_uppercase()))
}
fn write_font(m: &Map<String, Value>) -> KernelResult<String> {
    let mut s = String::from("<font>");
    if let Some(v) = m.get("fontFamily").and_then(Value::as_str) {
        s.push_str(&format!("<name val=\"{}\"/>", escape(v)));
    }
    if let Some(v) = m.get("fontSizePx").and_then(Value::as_f64) {
        s.push_str(&format!("<sz val=\"{}\"/>", v * 72.0 / 96.0));
    }
    for (key, tag) in [("bold", "b"), ("italic", "i"), ("strikethrough", "strike")] {
        if flag(m, key, false)? == 1 {
            s.push_str(&format!("<{tag}/>"));
        }
    }
    if flag(m, "underline", false)? == 1 {
        s.push_str(&format!(
            "<u val=\"{}\"/>",
            escape(
                m.get("underlineStyle")
                    .and_then(Value::as_str)
                    .unwrap_or("single")
            )
        ));
    }
    for key in ["superscript", "subscript"] {
        if flag(m, key, false)? == 1 {
            s.push_str(&format!("<vertAlign val=\"{key}\"/>"));
        }
    }
    if let Some(v) = m.get("textColor").and_then(Value::as_str) {
        s.push_str(&format!("<color rgb=\"{}\"/>", rgb(v)?));
    }
    if let Some(v) = m.get("fontTheme").and_then(Value::as_str) {
        s.push_str(&format!("<scheme val=\"{}\"/>", escape(v)));
    }
    s.push_str("</font>");
    Ok(s)
}
fn write_fill(m: &Map<String, Value>) -> KernelResult<String> {
    if let Some(f) = m.get("fill").and_then(Value::as_object) {
        if f.get("kind").and_then(Value::as_str) == Some("gradient") {
            let mut s = format!(
                "<fill><gradientFill type=\"{}\" degree=\"{}\">",
                escape(
                    f.get("gradientType")
                        .and_then(Value::as_str)
                        .unwrap_or("linear")
                ),
                f.get("degree").and_then(Value::as_f64).unwrap_or(0.0)
            );
            for stop in f
                .get("stops")
                .and_then(Value::as_array)
                .ok_or_else(|| error("CELL_STYLE_INVALID", "Gradient stops missing"))?
            {
                let position = stop
                    .get("position")
                    .and_then(Value::as_f64)
                    .ok_or_else(|| error("CELL_STYLE_INVALID", "Gradient position missing"))?;
                let color = stop
                    .get("color")
                    .and_then(Value::as_str)
                    .ok_or_else(|| error("CELL_STYLE_INVALID", "Gradient color missing"))?;
                s.push_str(&format!(
                    "<stop position=\"{position}\"><color rgb=\"{}\"/></stop>",
                    rgb(color)?
                ));
            }
            s.push_str("</gradientFill></fill>");
            return Ok(s);
        }
        let pattern = f.get("pattern").and_then(Value::as_str).unwrap_or("solid");
        let mut s = format!("<fill><patternFill patternType=\"{}\">", escape(pattern));
        for (key, tag) in [("foreground", "fgColor"), ("background", "bgColor")] {
            if let Some(v) = f.get(key).and_then(Value::as_str) {
                s.push_str(&format!("<{tag} rgb=\"{}\"/>", rgb(v)?));
            }
        }
        s.push_str("</patternFill></fill>");
        return Ok(s);
    }
    if let Some(v) = m.get("background").and_then(Value::as_str) {
        return Ok(format!("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/><bgColor indexed=\"64\"/></patternFill></fill>",rgb(v)?));
    }
    Ok("<fill><patternFill patternType=\"none\"/></fill>".into())
}
fn write_border(m: &Map<String, Value>) -> KernelResult<String> {
    let Some(b) = m.get("borders").and_then(Value::as_object) else {
        return Ok("<border/>".into());
    };
    let mut s = format!(
        "<border diagonalUp=\"{}\" diagonalDown=\"{}\">",
        flag(b, "diagonalUp", false)?,
        flag(b, "diagonalDown", false)?
    );
    for name in ["left", "right", "top", "bottom", "diagonal"] {
        if let Some(side) = b.get(name) {
            let style = side
                .get("style")
                .and_then(Value::as_str)
                .ok_or_else(|| error("CELL_STYLE_INVALID", "Border style missing"))?;
            let color = side
                .get("color")
                .and_then(Value::as_str)
                .ok_or_else(|| error("CELL_STYLE_INVALID", "Border color missing"))?;
            s.push_str(&format!(
                "<{name} style=\"{}\"><color rgb=\"{}\"/></{name}>",
                escape(style),
                rgb(color)?
            ));
        }
    }
    s.push_str("</border>");
    Ok(s)
}
fn write_alignment(m: &Map<String, Value>) -> KernelResult<String> {
    let mut s = String::from("<alignment");
    for (key, native) in [
        ("horizontalAlignment", "horizontal"),
        ("verticalAlignment", "vertical"),
    ] {
        if let Some(v) = m.get(key).and_then(Value::as_str) {
            s.push_str(&format!(" {native}=\"{}\"", escape(v)));
        }
    }
    for key in ["wrapText", "shrinkToFit"] {
        if m.contains_key(key) {
            s.push_str(&format!(" {key}=\"{}\"", flag(m, key, false)?));
        }
    }
    if let Some(v) = m.get("indent").and_then(Value::as_u64) {
        s.push_str(&format!(" indent=\"{v}\""));
    }
    if let Some(v) = m.get("textRotate").and_then(Value::as_i64) {
        if !(-90..=90).contains(&v) {
            return Err(error("CELL_STYLE_INVALID", "Rotation outside [-90,90]"));
        }
        s.push_str(&format!(
            " textRotation=\"{}\"",
            if v < 0 { 90 - v } else { v }
        ));
    } else if m.get("textOrientation").and_then(Value::as_str) == Some("stacked") {
        s.push_str(" textRotation=\"255\"");
    }
    if let Some(v) = m.get("readingOrder").and_then(Value::as_str) {
        let order = match v {
            "context" => 0,
            "ltr" => 1,
            "rtl" => 2,
            _ => return Err(error("CELL_STYLE_INVALID", "Invalid reading order")),
        };
        s.push_str(&format!(" readingOrder=\"{order}\""));
    }
    s.push_str("/>");
    Ok(s)
}
