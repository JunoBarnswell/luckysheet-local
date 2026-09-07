use crate::{Transaction, integer, invalid, text};
use kernel_core::*;
use serde_json::{Value, json};

pub(crate) fn apply(tx: &mut Transaction, id: &str, sheet: &str, p: &Value) -> KernelResult<bool> {
    if !matches!(
        id,
        "note.set"
            | "note.remove"
            | "note.visibility"
            | "comment.add"
            | "comment.reply"
            | "comment.reply.remove"
            | "comment.resolve"
            | "comment.remove"
            | "comment.update"
    ) {
        return Ok(false);
    }
    let mut review = tx
        .sheet(sheet)?
        .metadata
        .get("review")
        .cloned()
        .unwrap_or_else(
            || json!({"notesByCell":{},"notesById":{},"threadIdsByCell":{},"threadsById":{}}),
        );
    for field in ["notesByCell", "notesById", "threadIdsByCell", "threadsById"] {
        if !review[field].is_object() {
            return Err(invalid("Review index is malformed"));
        }
    }
    let (row, col) = if id.starts_with("note.") || id == "comment.add" {
        (integer(p, "row")?, integer(p, "column")?)
    } else {
        let thread = text(p, "threadId")?;
        let t = &review["threadsById"][thread];
        if !t.is_object() {
            return Err(KernelError::new("NOT_FOUND", "Comment thread not found"));
        }
        (integer(t, "row")?, integer(t, "column")?)
    };
    tx.read(sheet, row, col)?;
    let key = format!("{row}:{col}");
    if id.starts_with("note.") {
        let old = review["notesByCell"][&key].as_str().map(str::to_owned);
        if id == "note.set" {
            let note = p.get("note").ok_or_else(|| invalid("note is required"))?;
            let note_id = text(note, "id")?;
            if review["notesByCell"]
                .as_object()
                .unwrap()
                .iter()
                .any(|(k, v)| k != &key && v.as_str() == Some(note_id))
            {
                return Err(KernelError::new(
                    "CONFLICT",
                    "Note identity belongs to another cell",
                ));
            }
            if let Some(old) = old {
                if review["notesById"]
                    .as_object_mut()
                    .unwrap()
                    .remove(&old)
                    .is_none()
                {
                    return Err(invalid("Dangling note index"));
                }
            }
            review["notesByCell"][&key] = json!(note_id);
            review["notesById"][note_id] = note.clone();
        } else {
            let old = old.ok_or_else(|| KernelError::new("NOT_FOUND", "Note not found"))?;
            if !review["notesById"][&old].is_object() {
                return Err(invalid("Dangling note index"));
            }
            if id == "note.remove" {
                review["notesByCell"].as_object_mut().unwrap().remove(&key);
                review["notesById"].as_object_mut().unwrap().remove(&old);
            } else {
                if !p["visible"].is_boolean() {
                    return Err(invalid("visible must be Boolean"));
                }
                review["notesById"][&old]["visible"] = p["visible"].clone();
            }
        }
    } else if id == "comment.add" {
        let t = p
            .get("thread")
            .ok_or_else(|| invalid("thread is required"))?;
        let tid = text(t, "id")?;
        if text(t, "sheetId")? != sheet || integer(t, "row")? != row || integer(t, "column")? != col
        {
            return Err(invalid("Comment address differs from command"));
        }
        if !review["threadsById"][tid].is_null() {
            return Err(KernelError::new("CONFLICT", "Comment already exists"));
        }
        if review["threadIdsByCell"][&key].is_null() {
            review["threadIdsByCell"][&key] = json!([]);
        }
        review["threadIdsByCell"][&key]
            .as_array_mut()
            .ok_or_else(|| invalid("Comment cell index must be array"))?
            .push(json!(tid));
        review["threadsById"][tid] = t.clone();
    } else {
        let tid = text(p, "threadId")?;
        if id == "comment.remove" {
            let ids = review["threadIdsByCell"][&key]
                .as_array_mut()
                .ok_or_else(|| invalid("Dangling comment index"))?;
            if !ids.iter().any(|v| v.as_str() == Some(tid)) {
                return Err(invalid("Dangling comment index"));
            }
            ids.retain(|v| v.as_str() != Some(tid));
            if ids.is_empty() {
                review["threadIdsByCell"]
                    .as_object_mut()
                    .unwrap()
                    .remove(&key);
            }
            review["threadsById"].as_object_mut().unwrap().remove(tid);
        } else {
            let t = &mut review["threadsById"][tid];
            match id {
                "comment.reply" => {
                    let reply = p.get("reply").ok_or_else(|| invalid("reply is required"))?;
                    let rid = text(reply, "id")?;
                    if t["replies"].is_null() {
                        t["replies"] = json!([]);
                    }
                    let replies = t["replies"]
                        .as_array_mut()
                        .ok_or_else(|| invalid("Replies must be array"))?;
                    if replies.iter().any(|r| r["id"].as_str() == Some(rid)) {
                        return Err(KernelError::new("CONFLICT", "Reply already exists"));
                    }
                    replies.push(reply.clone());
                }
                "comment.reply.remove" => {
                    let rid = text(p, "replyId")?;
                    let replies = t["replies"]
                        .as_array_mut()
                        .ok_or_else(|| invalid("Replies must be array"))?;
                    let at = replies
                        .iter()
                        .position(|r| r["id"].as_str() == Some(rid))
                        .ok_or_else(|| KernelError::new("NOT_FOUND", "Reply not found"))?;
                    replies.remove(at);
                }
                "comment.resolve" => {
                    let resolved = p["resolved"]
                        .as_bool()
                        .ok_or_else(|| invalid("resolved must be Boolean"))?;
                    t["resolved"] = json!(resolved);
                    if resolved {
                        t["resolvedAt"] = json!(text(p, "resolvedAt")?);
                    } else {
                        t.as_object_mut().unwrap().remove("resolvedAt");
                    }
                }
                "comment.update" => {
                    if integer(p, "row")? != row
                        || integer(p, "column")? != col
                        || p.get("previousText") != t.get("text")
                    {
                        return Err(KernelError::new(
                            "CONFLICT",
                            "Comment changed before update",
                        ));
                    }
                    if !p["text"].is_string() {
                        return Err(invalid("Comment text must be text"));
                    }
                    t["text"] = p["text"].clone();
                }
                _ => unreachable!(),
            }
        }
    }
    tx.sheet_mut(sheet)?
        .metadata
        .insert("review".into(), review);
    tx.affected.push(RangeRef {
        sheet_id: sheet.into(),
        start_row: row,
        end_row: row,
        start_column: col,
        end_column: col,
    });
    Ok(true)
}
