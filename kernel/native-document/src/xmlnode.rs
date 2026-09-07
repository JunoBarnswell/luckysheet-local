//! Bounded metadata DOM. Payload worksheets and string dictionaries use streaming parsers.
use kernel_core::{KernelError, KernelResult};
use quick_xml::{events::Event, Reader};
use std::{collections::BTreeMap, ops::Range};

#[derive(Debug, Clone)]
pub(crate) struct Node {
    pub name: String,
    pub attributes: BTreeMap<String, String>,
    pub children: Vec<Node>,
    pub text: String,
    pub span: Range<usize>,
    pub opening_end: usize,
    pub closing_start: usize,
    pub empty: bool,
}
impl Node {
    pub fn local(&self) -> &str {
        self.name.rsplit(':').next().unwrap_or(&self.name)
    }
    pub fn attr(&self, key: &str) -> Option<&str> {
        self.attributes.get(key).map(String::as_str)
    }
    pub fn child(&self, key: &str) -> Option<&Node> {
        self.children.iter().find(|n| n.local() == key)
    }
    pub fn children_named<'a>(&'a self, key: &'a str) -> impl Iterator<Item = &'a Node> {
        self.children.iter().filter(move |n| n.local() == key)
    }
}
pub(crate) fn parse(bytes: &[u8], budget: u64) -> KernelResult<Node> {
    if bytes.len() as u64 > budget {
        return Err(error(
            "XML_METADATA_BUDGET",
            "Metadata XML exceeds its explicit memory budget",
        ));
    }
    let mut reader = Reader::from_reader(bytes);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    loop {
        let start = reader.buffer_position() as usize;
        match reader
            .read_event()
            .map_err(|e| error("XML_INVALID", e.to_string()))?
        {
            Event::Start(e) | Event::Empty(e) => {
                let end = reader.buffer_position() as usize;
                let empty = bytes.get(end.saturating_sub(2)..end) == Some(b"/>");
                nodes += 1;
                if nodes > 500_000 || stack.len() >= 128 {
                    return Err(error(
                        "XML_METADATA_BUDGET",
                        "Metadata XML node/depth budget exceeded",
                    ));
                }
                let name = String::from_utf8(e.name().as_ref().to_vec())
                    .map_err(|e| error("XML_INVALID", e.to_string()))?;
                let mut attributes = BTreeMap::new();
                for attribute in e.attributes() {
                    let attribute = attribute.map_err(|e| error("XML_INVALID", e.to_string()))?;
                    let key = String::from_utf8(attribute.key.as_ref().to_vec())
                        .map_err(|e| error("XML_INVALID", e.to_string()))?;
                    let value = attribute
                        .unescape_value()
                        .map_err(|e| error("XML_INVALID", e.to_string()))?
                        .into_owned();
                    if attributes.insert(key, value).is_some() {
                        return Err(error("XML_INVALID", "Duplicate XML attribute"));
                    }
                }
                let node = Node {
                    name,
                    attributes,
                    children: Vec::new(),
                    text: String::new(),
                    span: start..end,
                    opening_end: end,
                    closing_start: end,
                    empty,
                };
                if empty {
                    attach(&mut stack, &mut root, node)?;
                } else {
                    stack.push(node);
                }
            }
            Event::End(e) => {
                let mut node = stack
                    .pop()
                    .ok_or_else(|| error("XML_INVALID", "Unexpected closing element"))?;
                if node.name.as_bytes() != e.name().as_ref() {
                    return Err(error("XML_INVALID", "Mismatched closing element"));
                }
                node.closing_start = start;
                node.span.end = reader.buffer_position() as usize;
                attach(&mut stack, &mut root, node)?;
            }
            Event::Text(e) => {
                let value = e
                    .unescape()
                    .map_err(|e| error("XML_INVALID", e.to_string()))?;
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                }
            }
            Event::CData(e) => {
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(
                        &e.decode()
                            .map_err(|e| error("XML_INVALID", e.to_string()))?,
                    );
                }
            }
            Event::DocType(_) => {
                return Err(error(
                    "XML_DOCTYPE_FORBIDDEN",
                    "DTD/entity definitions are not allowed",
                ))
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err(error("XML_INVALID", "Unclosed element"));
    }
    root.ok_or_else(|| error("XML_INVALID", "XML has no root"))
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, node: Node) -> KernelResult<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(error("XML_INVALID", "Multiple root elements"));
    }
    Ok(())
}
pub(crate) fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\r', "&#13;")
}
pub(crate) fn error(code: &str, message: impl Into<String>) -> KernelError {
    KernelError::new(code, message).recover("Correct the named native document object and retry")
}
