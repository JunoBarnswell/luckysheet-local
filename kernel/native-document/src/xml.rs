use kernel_core::{KernelError, KernelResult};

pub(crate) fn attr<'a>(
    attrs: quick_xml::events::attributes::Attributes<'a>,
    name: &[u8],
) -> KernelResult<Option<String>> {
    for item in attrs {
        let item = item.map_err(|e| invalid(e.to_string()))?;
        if item.key.as_ref() == name {
            return Ok(Some(
                item.unescape_value()
                    .map_err(|e| invalid(e.to_string()))?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}
pub(crate) fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|b| *b == b':').next().unwrap_or(name)
}

/// Resolves an internal OPC relationship target against its source part.
/// Pass `None` for package-root relationships and otherwise the owning part
/// path, for example `xl/workbook.xml`.
pub(crate) fn resolve_relationship_target(
    target: &str,
    source_part: Option<&str>,
) -> KernelResult<String> {
    let target = target.trim();
    if target.is_empty() {
        return Err(target_error("relationship target is empty"));
    }
    let target = percent_decode(target)?;
    if target.contains('\\')
        || target.contains('?')
        || target.contains('#')
        || has_uri_scheme(&target)
    {
        return Err(target_error(target));
    }
    let base = if target.starts_with('/') {
        String::new()
    } else {
        let source = normalize_part_path(source_part.unwrap_or(""))?;
        source
            .rsplit_once('/')
            .map(|(dir, _)| dir.to_owned())
            .unwrap_or_default()
    };
    normalize_joined_path(&base, target.trim_start_matches('/'))
}
fn normalize_part_path(value: &str) -> KernelResult<String> {
    if value.is_empty() {
        Ok(String::new())
    } else {
        normalize_joined_path("", value.trim_start_matches('/'))
    }
}
fn normalize_joined_path(base: &str, target: &str) -> KernelResult<String> {
    let mut parts = Vec::new();
    let joined = if base.is_empty() {
        target.to_owned()
    } else {
        format!("{base}/{target}")
    };
    for part in joined.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                if parts.pop().is_none() {
                    return Err(target_error("relationship target escapes package root"));
                }
            }
            value if value.contains('\\') => return Err(target_error(value)),
            value => parts.push(value),
        }
    }
    if parts.is_empty() {
        return Err(target_error(target));
    }
    Ok(parts.join("/"))
}
fn has_uri_scheme(value: &str) -> bool {
    let Some(colon) = value.find(':') else {
        return false;
    };
    let Some(separator) = value.find('/') else {
        return colon > 0;
    };
    colon < separator
}
fn percent_decode(value: &str) -> KernelResult<String> {
    let bytes = value.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            if index + 2 >= bytes.len() {
                return Err(target_error("malformed percent escape"));
            }
            let high =
                hex(bytes[index + 1]).ok_or_else(|| target_error("malformed percent escape"))?;
            let low =
                hex(bytes[index + 2]).ok_or_else(|| target_error("malformed percent escape"))?;
            out.push((high << 4) | low);
            index += 3;
        } else {
            out.push(bytes[index]);
            index += 1;
        }
    }
    String::from_utf8(out).map_err(|e| target_error(e.to_string()))
}
fn hex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}
fn invalid(message: impl Into<String>) -> KernelError {
    KernelError::new("XML_INVALID", message).recover("Provide well-formed XML")
}
fn target_error(message: impl Into<String>) -> KernelError {
    KernelError::new("OOXML_RELATIONSHIP_TARGET_INVALID", message)
        .recover("Provide an internal OPC relationship target")
}
