//! Canonical source syntax for evaluation, editing and structural transformations.
//! Every span uses UTF-16 offsets, matching browser selection offsets.
use kernel_core::{CellAddress, KernelError, KernelResult, MAX_COLUMNS, MAX_ROWS};
use serde::{Deserialize, Serialize};

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Span {
    pub start: u32,
    pub end: u32,
}
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParsedCellReference {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub sheet_id: Option<String>,
    pub row: u32,
    pub column: u32,
    pub absolute_row: bool,
    pub absolute_column: bool,
}
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct Expr {
    #[serde(flatten)]
    pub node: Node,
    pub span: Span,
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub parenthesized: bool,
}
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase"
)]
pub enum Node {
    NumberLiteral {
        value: f64,
    },
    StringLiteral {
        value: String,
    },
    BooleanLiteral {
        value: bool,
    },
    CellReference {
        reference: ParsedCellReference,
    },
    InvalidReference {
        code: String,
    },
    ErrorLiteral {
        code: String,
    },
    NameReference {
        name: String,
    },
    RangeReference {
        start: Box<Expr>,
        end: Box<Expr>,
    },
    WholeColumnReference {
        #[serde(default, skip_serializing_if = "Option::is_none")]
        sheet_id: Option<String>,
        start_column: u32,
        end_column: u32,
        #[serde(default)]
        absolute_start: bool,
        #[serde(default)]
        absolute_end: bool,
    },
    WholeRowReference {
        #[serde(default, skip_serializing_if = "Option::is_none")]
        sheet_id: Option<String>,
        start_row: u32,
        end_row: u32,
        #[serde(default)]
        absolute_start: bool,
        #[serde(default)]
        absolute_end: bool,
    },
    UnaryExpression {
        operator: String,
        operand: Box<Expr>,
    },
    BinaryExpression {
        operator: String,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    FunctionCall {
        name: String,
        arguments: Vec<Expr>,
    },
    LambdaInvocation {
        callee: Box<Expr>,
        arguments: Vec<Expr>,
    },
    SpillReference {
        operand: Box<Expr>,
    },
    ArrayLiteral {
        rows: Vec<Vec<Expr>>,
    },
    MissingArgument {},
    ReferenceUnion {
        references: Vec<Expr>,
    },
    ReferenceIntersection {
        left: Box<Expr>,
        right: Box<Expr>,
    },
    TableReference {
        table_name: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        specifier: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        column_name: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        column_end_name: Option<String>,
        this_row: bool,
    },
}
fn fail(m: impl Into<String>) -> KernelError {
    KernelError::new("FORMULA_PARSE", m)
}
fn unsupported(m: &str) -> KernelError {
    KernelError::new("UNSUPPORTED_FEATURE", m)
}
fn make(node: Node, start: u32, end: u32) -> Expr {
    Expr {
        node,
        span: Span { start, end },
        parenthesized: false,
    }
}
#[derive(Clone, Debug, PartialEq)]
enum Kind {
    Word(String),
    Number(String),
    Text(String),
    Sheet(String),
    Structured(String),
    Error(String),
    Symbol(String),
    Space,
    End,
}
#[derive(Clone, Debug)]
struct Token {
    kind: Kind,
    start: u32,
    end: u32,
}
fn lex(source: &str) -> KernelResult<Vec<Token>> {
    if source.len() > 32768 {
        return Err(fail("Formula source budget exceeded (32768 bytes)"));
    }
    let chars: Vec<char> = source.chars().collect();
    let mut offsets = vec![0];
    for c in &chars {
        offsets.push(offsets.last().unwrap() + c.len_utf16() as u32);
    }
    let mut ts = Vec::new();
    let mut i = 0;
    let mut formula_marker = false;
    while i < chars.len() {
        let start = i;
        let c = chars[i];
        i += 1;
        let kind = if c.is_whitespace() {
            while i < chars.len() && chars[i].is_whitespace() {
                i += 1;
            }
            Kind::Space
        } else if c == '=' && !formula_marker && ts.iter().all(|t: &Token| t.kind == Kind::Space) {
            formula_marker = true;
            continue;
        } else if c == '"' || c == '\'' {
            let mut text = String::new();
            let mut closed = false;
            while i < chars.len() {
                let x = chars[i];
                i += 1;
                if x == c {
                    if i < chars.len() && chars[i] == c {
                        text.push(c);
                        i += 1;
                    } else {
                        closed = true;
                        break;
                    }
                } else {
                    text.push(x);
                }
            }
            if !closed {
                return Err(fail("Unterminated quoted literal"));
            }
            if c == '"' {
                Kind::Text(text)
            } else {
                Kind::Sheet(text)
            }
        } else if c == '[' {
            let mut depth = 1;
            while i < chars.len() && depth > 0 {
                if chars[i] == '\'' && i + 1 < chars.len() {
                    i += 2;
                    continue;
                }
                if chars[i] == '[' {
                    depth += 1;
                }
                if chars[i] == ']' {
                    depth -= 1;
                }
                i += 1;
            }
            if depth != 0 {
                return Err(fail("Unterminated structured reference"));
            }
            Kind::Structured(chars[start..i].iter().collect())
        } else if c == '#' {
            let rest: String = chars[start..].iter().collect();
            let upper = rest.to_ascii_uppercase();
            let known = [
                "#GETTING_DATA",
                "#BLOCKED!",
                "#CONNECT!",
                "#UNKNOWN!",
                "#PYTHON!",
                "#SPILL!",
                "#VALUE!",
                "#DIV/0!",
                "#FIELD!",
                "#CALC!",
                "#NULL!",
                "#NAME?",
                "#BUSY!",
                "#REF!",
                "#NUM!",
                "#N/A",
            ];
            if let Some(e) = known.iter().find(|e| upper.starts_with(**e)) {
                i = start + e.len();
                Kind::Error(e.to_string())
            } else {
                Kind::Symbol("#".into())
            }
        } else if c.is_ascii_digit() || (c == '.' && i < chars.len() && chars[i].is_ascii_digit()) {
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
            if c != '.' && i < chars.len() && chars[i] == '.' {
                i += 1;
                while i < chars.len() && chars[i].is_ascii_digit() {
                    i += 1;
                }
            }
            if i < chars.len() && matches!(chars[i], 'e' | 'E') {
                i += 1;
                if i < chars.len() && matches!(chars[i], '+' | '-') {
                    i += 1;
                }
                let e = i;
                while i < chars.len() && chars[i].is_ascii_digit() {
                    i += 1;
                }
                if i == e {
                    return Err(fail("Exponent requires digits"));
                }
            }
            Kind::Number(chars[start..i].iter().collect())
        } else if c.is_alphabetic() || c == '_' || c == '$' || c == '\\' {
            while i < chars.len()
                && (chars[i].is_alphanumeric() || matches!(chars[i], '_' | '$' | '.' | '\\'))
            {
                i += 1;
            }
            Kind::Word(chars[start..i].iter().collect())
        } else if "()+-*/^&=<>%,;:{}!@".contains(c) {
            let mut s = c.to_string();
            if i < chars.len() && matches!((c, chars[i]), ('<', '=') | ('>', '=') | ('<', '>')) {
                s.push(chars[i]);
                i += 1;
            }
            Kind::Symbol(s)
        } else {
            return Err(fail(format!("Invalid formula character {c}")));
        };
        ts.push(Token {
            kind,
            start: offsets[start],
            end: offsets[i],
        });
        if ts.len() > 1024 {
            return Err(fail("Formula token budget exceeded (1024 tokens)"));
        }
    }
    let end = *offsets.last().unwrap();
    ts.push(Token {
        kind: Kind::End,
        start: end,
        end,
    });
    Ok(ts)
}
struct Parser {
    tokens: Vec<Token>,
    index: usize,
    depth: usize,
}
impl Parser {
    fn skip(&mut self) {
        while self.tokens[self.index].kind == Kind::Space {
            self.index += 1;
        }
    }
    fn peek(&mut self) -> Kind {
        let mut index = self.index;
        while self.tokens[index].kind == Kind::Space {
            index += 1;
        }
        self.tokens[index].kind.clone()
    }
    fn take(&mut self) -> Token {
        self.skip();
        let t = self.tokens[self.index].clone();
        if t.kind != Kind::End {
            self.index += 1;
        }
        t
    }
    fn eat(&mut self, s: &str) -> bool {
        if self.peek() == Kind::Symbol(s.into()) {
            self.skip();
            self.index += 1;
            true
        } else {
            false
        }
    }
    fn end(&self) -> u32 {
        self.tokens[self.index.saturating_sub(1)].end
    }
    fn expression(&mut self, min: u8, union: bool) -> KernelResult<Expr> {
        self.depth += 1;
        if self.depth > 96 {
            return Err(fail("Formula nesting budget exceeded (96 levels)"));
        }
        let result = self.expression_inner(min, union);
        self.depth -= 1;
        result
    }
    fn expression_inner(&mut self, min: u8, union: bool) -> KernelResult<Expr> {
        let token = self.take();
        let start = token.start;
        let mut left = match token.kind {
            Kind::Symbol(ref s) if matches!(s.as_str(), "+" | "-" | "@") => {
                let operand = self.expression(7, union)?;
                let end = operand.span.end;
                make(
                    Node::UnaryExpression {
                        operator: s.clone(),
                        operand: Box::new(operand),
                    },
                    start,
                    end,
                )
            }
            Kind::Symbol(ref s) if s == "(" => {
                let mut e = self.expression(0, true)?;
                if !self.eat(")") {
                    return Err(fail("Expected closing parenthesis"));
                }
                e.span = Span {
                    start,
                    end: self.end(),
                };
                e.parenthesized = true;
                e
            }
            Kind::Symbol(ref s) if s == "{" => self.array(start)?,
            Kind::Number(ref v) => {
                if self.peek() == Kind::Symbol(":".into()) {
                    self.reference_atom(None, v.clone(), start, token.end)?
                } else {
                    let value = v.parse::<f64>().map_err(|_| fail("Invalid number"))?;
                    if !value.is_finite() {
                        return Err(fail("Non-finite number literal"));
                    }
                    make(Node::NumberLiteral { value }, start, token.end)
                }
            }
            Kind::Text(value) => make(Node::StringLiteral { value }, start, token.end),
            Kind::Error(code) => make(
                if code == "#REF!" {
                    Node::InvalidReference { code }
                } else {
                    Node::ErrorLiteral { code }
                },
                start,
                token.end,
            ),
            Kind::Word(word) => self.word(word, start, token.end)?,
            Kind::Sheet(sheet) => {
                if !self.eat("!") {
                    return Err(fail("Quoted worksheet requires !"));
                }
                let t = self.take();
                match t.kind {
                    Kind::Word(v) | Kind::Number(v) => {
                        self.reference_atom(Some(sheet), v, start, t.end)?
                    }
                    _ => return Err(fail("Expected worksheet reference")),
                }
            }
            Kind::Structured(_) => {
                return Err(unsupported(
                    "External or unqualified structured reference requires explicit workbook/table context",
                ));
            }
            _ => return Err(fail("Expected formula expression")),
        };
        loop {
            let had_space = self.tokens[self.index].kind == Kind::Space;
            let next = self.peek();
            if self.eat("%") {
                let end = self.end();
                left = make(
                    Node::UnaryExpression {
                        operator: "%".into(),
                        operand: Box::new(left),
                    },
                    start,
                    end,
                );
                continue;
            }
            if self.eat("#") {
                let end = self.end();
                left = make(
                    Node::SpillReference {
                        operand: Box::new(left),
                    },
                    start,
                    end,
                );
                continue;
            }
            if next == Kind::Symbol("(".into()) && !had_space {
                let arguments = self.arguments()?;
                left = make(
                    Node::LambdaInvocation {
                        callee: Box::new(left),
                        arguments,
                    },
                    start,
                    self.end(),
                );
                continue;
            }
            let (op, prec) = match &next {
                Kind::Symbol(s) => match s.as_str() {
                    "," if union => (s.clone(), 0),
                    "=" | "<>" | "<" | "<=" | ">" | ">=" => (s.clone(), 1),
                    "&" => (s.clone(), 2),
                    "+" | "-" => (s.clone(), 3),
                    "*" | "/" => (s.clone(), 4),
                    "^" => (s.clone(), 5),
                    ":" => (s.clone(), 9),
                    _ => break,
                },
                _ if had_space
                    && is_ref(&left)
                    && matches!(next, Kind::Word(_) | Kind::Sheet(_)) =>
                {
                    (" ".into(), 8)
                }
                _ => break,
            };
            if prec < min {
                break;
            }
            if op != " " {
                self.take();
            }
            // Excel exponentiation is left associative; unary signs bind more tightly.
            let right = self.expression(prec + 1, union)?;
            let end = right.span.end;
            left = if op == ":" {
                range(left, right, start, end)?
            } else if op == "," {
                if !is_ref(&left) || !is_ref(&right) {
                    return Err(fail("Union operands must be references"));
                }
                make(
                    Node::ReferenceUnion {
                        references: vec![left, right],
                    },
                    start,
                    end,
                )
            } else if op == " " {
                if !is_ref(&right) {
                    return Err(fail("Intersection operands must be references"));
                }
                make(
                    Node::ReferenceIntersection {
                        left: Box::new(left),
                        right: Box::new(right),
                    },
                    start,
                    end,
                )
            } else {
                make(
                    Node::BinaryExpression {
                        operator: op,
                        left: Box::new(left),
                        right: Box::new(right),
                    },
                    start,
                    end,
                )
            };
        }
        Ok(left)
    }
    fn word(&mut self, word: String, start: u32, end: u32) -> KernelResult<Expr> {
        let following: Vec<&Kind> = self.tokens[self.index..]
            .iter()
            .filter(|t| t.kind != Kind::Space)
            .take(3)
            .map(|t| &t.kind)
            .collect();
        if matches!(following.as_slice(), [Kind::Symbol(colon), Kind::Word(_) | Kind::Sheet(_), Kind::Symbol(bang)] if colon == ":" && bang == "!")
        {
            return Err(unsupported(
                "3D references require worksheet-range ownership, which is not implemented",
            ));
        }
        if self.peek() == Kind::Symbol("(".into()) {
            let arguments = self.arguments()?;
            return Ok(make(
                Node::FunctionCall {
                    name: word.to_ascii_uppercase(),
                    arguments,
                },
                start,
                self.end(),
            ));
        }
        if self.eat("!") {
            let t = self.take();
            return match t.kind {
                Kind::Word(v) | Kind::Number(v) => self.reference_atom(Some(word), v, start, t.end),
                _ => Err(fail("Expected worksheet reference")),
            };
        }
        if let Kind::Structured(raw) = self.peek() {
            self.take();
            return structured(word, raw, start, self.end());
        }
        if word.eq_ignore_ascii_case("TRUE") || word.eq_ignore_ascii_case("FALSE") {
            return Ok(make(
                Node::BooleanLiteral {
                    value: word.eq_ignore_ascii_case("TRUE"),
                },
                start,
                end,
            ));
        }
        if self.peek() == Kind::Symbol(":".into()) || looks_cell(&word) {
            return self.reference_atom(None, word, start, end);
        }
        Ok(make(Node::NameReference { name: word }, start, end))
    }
    fn reference_atom(
        &mut self,
        sheet: Option<String>,
        word: String,
        start: u32,
        end: u32,
    ) -> KernelResult<Expr> {
        if sheet.as_ref().is_some_and(|s| s.contains(':')) {
            return Err(unsupported(
                "3D references require worksheet-range ownership, which is not implemented",
            ));
        }
        if sheet
            .as_ref()
            .is_some_and(|s| s.contains('[') || s.contains(']'))
        {
            return Err(unsupported(
                "External references require external-workbook ownership, which is not implemented",
            ));
        }
        let a = endpoint(&word, sheet.clone())?;
        if self.eat(":") {
            let t = self.take();
            let endpoint_start = t.start;
            let mut end_sheet = None;
            let mut text = match t.kind {
                Kind::Word(v) | Kind::Number(v) | Kind::Sheet(v) => v,
                _ => return Err(fail("Expected range endpoint")),
            };
            if self.eat("!") {
                end_sheet = Some(text);
                let t = self.take();
                text = match t.kind {
                    Kind::Word(v) | Kind::Number(v) => v,
                    _ => return Err(fail("Expected qualified range endpoint")),
                };
            }
            if end_sheet
                .as_ref()
                .is_some_and(|s| sheet.as_ref().is_some_and(|a| !s.eq_ignore_ascii_case(a)))
            {
                return Err(unsupported(
                    "Cross-sheet range endpoints require a 3D reference",
                ));
            }
            let mut b = endpoint(&text, end_sheet)?;
            if a.1 != b.1 {
                return Err(fail("Range endpoints must have the same domain"));
            }
            let a_span = end;
            let end = self.end();
            return Ok(match a.1 {
                1 => make(
                    Node::WholeColumnReference {
                        sheet_id: sheet,
                        start_column: a.0.column,
                        end_column: b.0.column,
                        absolute_start: a.0.absolute_column,
                        absolute_end: b.0.absolute_column,
                    },
                    start,
                    end,
                ),
                2 => make(
                    Node::WholeRowReference {
                        sheet_id: sheet,
                        start_row: a.0.row,
                        end_row: b.0.row,
                        absolute_start: a.0.absolute_row,
                        absolute_end: b.0.absolute_row,
                    },
                    start,
                    end,
                ),
                _ => {
                    let b_start = endpoint_start;
                    if b.0.sheet_id == sheet {
                        b.0.sheet_id = None;
                    }
                    make(
                        Node::RangeReference {
                            start: Box::new(make(
                                Node::CellReference { reference: a.0 },
                                start,
                                a_span,
                            )),
                            end: Box::new(make(
                                Node::CellReference { reference: b.0 },
                                b_start,
                                end,
                            )),
                        },
                        start,
                        end,
                    )
                }
            });
        }
        if a.1 != 0 {
            return Err(fail("Whole row/column references require a range"));
        }
        Ok(make(Node::CellReference { reference: a.0 }, start, end))
    }
    fn arguments(&mut self) -> KernelResult<Vec<Expr>> {
        self.eat("(");
        let mut args = Vec::new();
        if self.eat(")") {
            return Ok(args);
        }
        loop {
            if matches!(self.peek(),Kind::Symbol(ref s) if s==","||s==")") {
                let at = self.tokens[self.index].start;
                args.push(make(Node::MissingArgument {}, at, at));
            } else {
                args.push(self.expression(0, false)?);
            }
            if self.eat(")") {
                return Ok(args);
            }
            if !self.eat(",") {
                return Err(fail("Expected argument separator or closing parenthesis"));
            }
        }
    }
    fn array(&mut self, start: u32) -> KernelResult<Expr> {
        let mut rows = Vec::new();
        let mut row = Vec::new();
        if self.eat("}") {
            return Err(fail("Array constant cannot be empty"));
        }
        loop {
            if matches!(self.peek(),Kind::Symbol(ref s) if s==","||s==";"||s=="}") {
                let at = self.tokens[self.index].start;
                row.push(make(Node::MissingArgument {}, at, at));
            } else {
                row.push(self.expression(0, false)?);
            }
            if self.eat(",") {
                continue;
            }
            rows.push(row);
            row = Vec::new();
            if self.eat("}") {
                break;
            }
            if !self.eat(";") {
                return Err(fail("Expected array separator"));
            }
        }
        if rows.iter().any(|r| r.len() != rows[0].len()) {
            return Err(fail("Array rows must have equal widths"));
        }
        Ok(make(Node::ArrayLiteral { rows }, start, self.end()))
    }
}
fn looks_cell(word: &str) -> bool {
    let s = word.trim_start_matches('$');
    let n = s.bytes().take_while(|c| c.is_ascii_alphabetic()).count();
    n > 0
        && s[n..]
            .trim_start_matches('$')
            .bytes()
            .all(|c| c.is_ascii_digit())
        && !s[n..].trim_start_matches('$').is_empty()
}
// domain: 0 cell, 1 column, 2 row. This is the sole A1 anchor parser.
pub(crate) fn endpoint(
    word: &str,
    sheet_id: Option<String>,
) -> KernelResult<(ParsedCellReference, u8)> {
    let mut s = word;
    let first = s.starts_with('$');
    if first {
        s = &s[1..];
    }
    let n = s.bytes().take_while(|c| c.is_ascii_alphabetic()).count();
    let col = &s[..n];
    s = &s[n..];
    let absolute_row = if col.is_empty() {
        first
    } else {
        let b = s.starts_with('$');
        if b {
            s = &s[1..];
        }
        b
    };
    if col.is_empty() && s.is_empty() || !s.bytes().all(|c| c.is_ascii_digit()) {
        return Err(fail("Invalid A1 reference"));
    }
    let domain = if col.is_empty() {
        2
    } else if s.is_empty() {
        1
    } else {
        0
    };
    let mut column = 0u32;
    for c in col.bytes() {
        column = column
            .checked_mul(26)
            .and_then(|n| n.checked_add((c.to_ascii_uppercase() - b'A' + 1) as u32))
            .ok_or_else(|| fail("Column outside worksheet bounds"))?;
    }
    if !col.is_empty() && (column == 0 || column > MAX_COLUMNS) {
        return Err(fail("Column outside worksheet bounds"));
    }
    let row = if s.is_empty() {
        0
    } else {
        let r = s.parse::<u32>().map_err(|_| fail("Invalid row"))?;
        if r == 0 || r > MAX_ROWS {
            return Err(fail("Row outside worksheet bounds"));
        }
        r - 1
    };
    Ok((
        ParsedCellReference {
            sheet_id,
            row,
            column: column.saturating_sub(1),
            absolute_row,
            absolute_column: !col.is_empty() && first,
        },
        domain,
    ))
}
fn range(a: Expr, b: Expr, start: u32, end: u32) -> KernelResult<Expr> {
    if !matches!(a.node, Node::CellReference { .. })
        || !matches!(b.node, Node::CellReference { .. })
    {
        return Err(fail("Range operands must be cell references"));
    }
    Ok(make(
        Node::RangeReference {
            start: Box::new(a),
            end: Box::new(b),
        },
        start,
        end,
    ))
}
fn is_ref(e: &Expr) -> bool {
    matches!(
        e.node,
        Node::CellReference { .. }
            | Node::RangeReference { .. }
            | Node::WholeColumnReference { .. }
            | Node::WholeRowReference { .. }
            | Node::NameReference { .. }
            | Node::TableReference { .. }
            | Node::ReferenceUnion { .. }
            | Node::ReferenceIntersection { .. }
            | Node::SpillReference { .. }
    )
}
fn structured(table_name: String, raw: String, start: u32, end: u32) -> KernelResult<Expr> {
    let mut inner = &raw[1..raw.len() - 1];
    let mut specifier = None;
    let mut this_row = false;
    if inner.starts_with("@[") {
        this_row = true;
        inner = &inner[1..];
    }
    let mut columns = Vec::new();
    // Structured fields are lexical identifiers; escapes are decoded after
    // deciding whether an unescaped leading #/@ is a selector.
    let mut column_range = false;
    let fields: Vec<String> = if inner.starts_with('[') {
        let mut fields = Vec::new();
        let mut field = String::new();
        let mut in_field = false;
        let mut escaped = false;
        for c in inner.chars() {
            if escaped {
                field.push(c);
                escaped = false;
                continue;
            }
            if c == '\'' && in_field {
                field.push(c);
                escaped = true;
                continue;
            }
            match c {
                '[' if !in_field => {
                    in_field = true;
                    field.clear();
                }
                ']' if in_field => {
                    fields.push(field.clone());
                    in_field = false;
                }
                ':' if !in_field => column_range = true,
                ',' | ' ' if !in_field => {}
                _ if in_field => field.push(c),
                _ => return Err(fail("Invalid structured reference separator")),
            }
        }
        if in_field || escaped {
            return Err(fail("Unterminated structured field"));
        }
        fields
    } else {
        vec![inner.to_string()]
    };
    for mut field in fields {
        if field.starts_with('@') {
            this_row = true;
            field = field[1..].to_string();
        }
        if field.starts_with('#') {
            let value = match field.to_ascii_lowercase().as_str() {
                "#all" => Some("all"),
                "#headers" => Some("headers"),
                "#data" => Some("data"),
                "#totals" => Some("totals"),
                "#this row" => {
                    this_row = true;
                    None
                }
                _ => return Err(unsupported("Unsupported table reference specifier")),
            };
            if let Some(value) = value {
                if specifier.is_some() {
                    return Err(unsupported(
                        "Multiple structured row selectors cannot be represented by the canonical table contract",
                    ));
                }
                specifier = Some(value.into());
            }
        } else if !field.is_empty() {
            let mut decoded = String::new();
            let mut chars = field.chars();
            while let Some(c) = chars.next() {
                if c == '\'' {
                    decoded.push(
                        chars
                            .next()
                            .ok_or_else(|| fail("Unterminated structured field escape"))?,
                    );
                } else {
                    decoded.push(c);
                }
            }
            columns.push(decoded);
        }
    }
    if columns.len() > 2 || columns.len() == 2 && !column_range {
        return Err(unsupported(
            "Non-contiguous structured columns are unsupported",
        ));
    }
    if columns.is_empty() && specifier.is_none() && !this_row {
        return Err(fail("Structured reference requires a selector or column"));
    }
    Ok(make(
        Node::TableReference {
            table_name,
            specifier,
            column_name: columns.first().cloned(),
            column_end_name: columns.get(1).cloned(),
            this_row,
        },
        start,
        end,
    ))
}
pub fn parse_editor(source: &str, current: &CellAddress) -> KernelResult<Expr> {
    current.validate()?;
    parse_source(source)
}
pub fn parse_source(source: &str) -> KernelResult<Expr> {
    let mut p = Parser {
        tokens: lex(source)?,
        index: 0,
        depth: 0,
    };
    let e = p.expression(0, true)?;
    if p.peek() != Kind::End {
        return Err(fail("Unexpected token after formula"));
    }
    validate_editor(&e)?;
    Ok(e)
}
pub(crate) fn children(e: &Expr) -> Vec<&Expr> {
    match &e.node {
        Node::RangeReference { start, end } => vec![start, end],
        Node::UnaryExpression { operand, .. } | Node::SpillReference { operand } => vec![operand],
        Node::BinaryExpression { left, right, .. }
        | Node::ReferenceIntersection { left, right } => vec![left, right],
        Node::FunctionCall { arguments, .. } => arguments.iter().collect(),
        Node::LambdaInvocation { callee, arguments } => std::iter::once(callee.as_ref())
            .chain(arguments.iter())
            .collect(),
        Node::ArrayLiteral { rows } => rows.iter().flatten().collect(),
        Node::ReferenceUnion { references } => references.iter().collect(),
        _ => vec![],
    }
}
pub(crate) fn map_children(
    e: &Expr,
    f: &mut impl FnMut(&Expr) -> KernelResult<Expr>,
) -> KernelResult<Expr> {
    let mut out = e.clone();
    match &mut out.node {
        Node::RangeReference { start, end } => {
            **start = f(start)?;
            **end = f(end)?;
        }
        Node::UnaryExpression { operand, .. } | Node::SpillReference { operand } => {
            **operand = f(operand)?
        }
        Node::BinaryExpression { left, right, .. }
        | Node::ReferenceIntersection { left, right } => {
            **left = f(left)?;
            **right = f(right)?;
        }
        Node::FunctionCall { arguments, .. } => {
            for a in arguments {
                *a = f(a)?;
            }
        }
        Node::LambdaInvocation { callee, arguments } => {
            **callee = f(callee)?;
            for a in arguments {
                *a = f(a)?;
            }
        }
        Node::ArrayLiteral { rows } => {
            for a in rows.iter_mut().flatten() {
                *a = f(a)?;
            }
        }
        Node::ReferenceUnion { references } => {
            for a in references {
                *a = f(a)?;
            }
        }
        _ => {}
    }
    Ok(out)
}
pub(crate) fn column_label(mut column: u32) -> String {
    let mut s = String::new();
    column += 1;
    while column > 0 {
        s.push((b'A' + ((column - 1) % 26) as u8) as char);
        column = (column - 1) / 26;
    }
    s.chars().rev().collect()
}
pub(crate) fn sheet_prefix(s: &Option<String>) -> String {
    s.as_ref()
        .map(|s| {
            if !s.is_empty()
                && s.chars()
                    .all(|c| c.is_alphanumeric() || c == '_' || c == '.')
                && !s.chars().next().unwrap().is_ascii_digit()
            {
                format!("{s}!")
            } else {
                format!("'{}'!", s.replace('\'', "''"))
            }
        })
        .unwrap_or_default()
}
pub(crate) fn format_ref(r: &ParsedCellReference) -> String {
    format!(
        "{}{}{}{}{}",
        sheet_prefix(&r.sheet_id),
        if r.absolute_column { "$" } else { "" },
        column_label(r.column),
        if r.absolute_row { "$" } else { "" },
        r.row + 1
    )
}
fn precedence(e: &Expr) -> u8 {
    match &e.node {
        Node::ReferenceUnion { .. } => 0,
        Node::BinaryExpression { operator, .. } => match operator.as_str() {
            "=" | "<>" | "<" | "<=" | ">" | ">=" => 1,
            "&" => 2,
            "+" | "-" => 3,
            "*" | "/" => 4,
            "^" => 5,
            _ => 1,
        },
        Node::UnaryExpression { .. } => 7,
        Node::ReferenceIntersection { .. } => 8,
        _ => 10,
    }
}
fn format_at(e: &Expr, min: u8) -> String {
    let text = format_editor_inner(e);
    if e.parenthesized || precedence(e) < min {
        format!("({text})")
    } else {
        text
    }
}
pub fn format_editor(e: &Expr) -> String {
    format_at(e, 0)
}
fn format_editor_inner(e: &Expr) -> String {
    match &e.node {
        Node::NumberLiteral { value } => value.to_string(),
        Node::StringLiteral { value } => format!("\"{}\"", value.replace('"', "\"\"")),
        Node::BooleanLiteral { value } => {
            if *value {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        Node::CellReference { reference } => format_ref(reference),
        Node::NameReference { name } => name.clone(),
        Node::InvalidReference { code } | Node::ErrorLiteral { code } => code.clone(),
        Node::MissingArgument {} => String::new(),
        Node::RangeReference { start, end } => {
            format!("{}:{}", format_editor(start), format_editor(end))
        }
        Node::WholeColumnReference {
            sheet_id,
            start_column,
            end_column,
            absolute_start,
            absolute_end,
        } => format!(
            "{}{}{}:{}{}",
            sheet_prefix(sheet_id),
            if *absolute_start { "$" } else { "" },
            column_label(*start_column),
            if *absolute_end { "$" } else { "" },
            column_label(*end_column)
        ),
        Node::WholeRowReference {
            sheet_id,
            start_row,
            end_row,
            absolute_start,
            absolute_end,
        } => format!(
            "{}{}{}:{}{}",
            sheet_prefix(sheet_id),
            if *absolute_start { "$" } else { "" },
            start_row + 1,
            if *absolute_end { "$" } else { "" },
            end_row + 1
        ),
        Node::UnaryExpression { operator, operand } => {
            if operator == "%" {
                format!("{}%", format_at(operand, 7))
            } else {
                format!("{}{}", operator, format_at(operand, 7))
            }
        }
        Node::BinaryExpression {
            operator,
            left,
            right,
        } => {
            let p = precedence(e);
            format!(
                "{}{}{}",
                format_at(left, p),
                operator,
                format_at(right, p + 1)
            )
        }
        Node::ReferenceIntersection { left, right } => {
            format!("{} {}", format_at(left, 8), format_at(right, 9))
        }
        Node::ReferenceUnion { references } => references
            .iter()
            .map(format_editor)
            .collect::<Vec<_>>()
            .join(","),
        Node::FunctionCall { name, arguments } => format!(
            "{}({})",
            name,
            arguments
                .iter()
                .map(|e| format_at(e, 1))
                .collect::<Vec<_>>()
                .join(",")
        ),
        Node::LambdaInvocation { callee, arguments } => format!(
            "{}({})",
            format_at(callee, 10),
            arguments
                .iter()
                .map(|e| format_at(e, 1))
                .collect::<Vec<_>>()
                .join(",")
        ),
        Node::SpillReference { operand } => format!("{}#", format_at(operand, 10)),
        Node::ArrayLiteral { rows } => format!(
            "{{{}}}",
            rows.iter()
                .map(|row| row
                    .iter()
                    .map(|e| format_at(e, 1))
                    .collect::<Vec<_>>()
                    .join(","))
                .collect::<Vec<_>>()
                .join(";")
        ),
        Node::TableReference {
            table_name,
            specifier,
            column_name,
            column_end_name,
            this_row,
        } => {
            let esc = |s: &str| {
                s.chars()
                    .flat_map(|c| {
                        if matches!(c, '[' | ']' | '#' | '\'' | '@') {
                            vec!['\'', c]
                        } else {
                            vec![c]
                        }
                    })
                    .collect::<String>()
            };
            let mut fields = Vec::new();
            if let Some(s) = specifier {
                fields.push(format!(
                    "[#{}]",
                    match s.as_str() {
                        "all" => "All",
                        "headers" => "Headers",
                        "totals" => "Totals",
                        _ => "Data",
                    }
                ));
            }
            if *this_row {
                fields.push("[#This Row]".into());
            }
            if let Some(c) = column_name {
                let mut col = format!("[{}]", esc(c));
                if let Some(end) = column_end_name {
                    col.push_str(&format!(":[{}]", esc(end)));
                }
                fields.push(col);
            }
            if fields.len() == 1 && specifier.is_none() && !*this_row && column_end_name.is_none() {
                format!("{}{}", table_name, fields[0])
            } else {
                format!("{}[{}]", table_name, fields.join(","))
            }
        }
    }
}
pub(crate) fn invalid(e: &Expr) -> Expr {
    let mut out = e.clone();
    out.node = Node::InvalidReference {
        code: "#REF!".into(),
    };
    out
}
fn offset_coord(value: u32, delta: i32, absolute: bool, limit: u32) -> Option<u32> {
    let n = value as i64 + if absolute { 0 } else { delta as i64 };
    (n >= 0 && n < limit as i64).then_some(n as u32)
}
pub fn offset_editor(e: &Expr, dr: i32, dc: i32) -> Expr {
    let mut out =
        map_children(e, &mut |x| Ok(offset_editor(x, dr, dc))).expect("infallible traversal");
    match &mut out.node {
        Node::CellReference { reference: r } => {
            let row = offset_coord(r.row, dr, r.absolute_row, MAX_ROWS);
            let col = offset_coord(r.column, dc, r.absolute_column, MAX_COLUMNS);
            if let (Some(row), Some(column)) = (row, col) {
                r.row = row;
                r.column = column;
            } else {
                return invalid(e);
            }
        }
        Node::RangeReference { start, end }
            if matches!(start.node, Node::InvalidReference { .. })
                || matches!(end.node, Node::InvalidReference { .. }) =>
        {
            return invalid(e);
        }
        Node::WholeColumnReference {
            start_column,
            end_column,
            absolute_start,
            absolute_end,
            ..
        } => {
            if let (Some(a), Some(b)) = (
                offset_coord(*start_column, dc, *absolute_start, MAX_COLUMNS),
                offset_coord(*end_column, dc, *absolute_end, MAX_COLUMNS),
            ) {
                *start_column = a;
                *end_column = b;
            } else {
                return invalid(e);
            }
        }
        Node::WholeRowReference {
            start_row,
            end_row,
            absolute_start,
            absolute_end,
            ..
        } => {
            if let (Some(a), Some(b)) = (
                offset_coord(*start_row, dr, *absolute_start, MAX_ROWS),
                offset_coord(*end_row, dr, *absolute_end, MAX_ROWS),
            ) {
                *start_row = a;
                *end_row = b;
            } else {
                return invalid(e);
            }
        }
        _ => {}
    }
    out
}
pub fn references_editor(e: &Expr) -> Vec<ParsedCellReference> {
    let mut out = Vec::new();
    fn visit(e: &Expr, inherited: Option<&String>, out: &mut Vec<ParsedCellReference>) {
        match &e.node {
            Node::CellReference { reference } => {
                let mut r = reference.clone();
                if r.sheet_id.is_none() {
                    r.sheet_id = inherited.cloned();
                }
                out.push(r);
            }
            Node::RangeReference { start, end } => {
                visit(start, inherited, out);
                let sheet = match &start.node {
                    Node::CellReference { reference } => reference.sheet_id.as_ref().or(inherited),
                    _ => inherited,
                };
                visit(end, sheet, out);
            }
            _ => {
                for c in children(e) {
                    visit(c, inherited, out);
                }
            }
        }
    }
    visit(e, None, &mut out);
    out
}
pub fn remap_editor(
    e: &Expr,
    mappings: &[(ParsedCellReference, Option<ParsedCellReference>)],
) -> Expr {
    fn visit(
        e: &Expr,
        mappings: &[(ParsedCellReference, Option<ParsedCellReference>)],
        inherited: Option<&String>,
    ) -> Expr {
        if let Node::CellReference { reference } = &e.node {
            let mut resolved = reference.clone();
            if resolved.sheet_id.is_none() {
                resolved.sheet_id = inherited.cloned();
            }
            if let Some((_, to)) = mappings.iter().find(|(from, _)| from == &resolved) {
                let Some(mut r) = to.clone() else {
                    return invalid(e);
                };
                if r.row >= MAX_ROWS || r.column >= MAX_COLUMNS {
                    return invalid(e);
                }
                if reference.sheet_id.is_none() && r.sheet_id.as_ref() == inherited {
                    r.sheet_id = None;
                }
                let mut out = e.clone();
                out.node = Node::CellReference { reference: r };
                return out;
            }
            return e.clone();
        }
        if let Node::RangeReference { start, end } = &e.node {
            let sheet = match &start.node {
                Node::CellReference { reference } => reference.sheet_id.as_ref().or(inherited),
                _ => inherited,
            };
            let a = visit(start, mappings, inherited);
            let b = visit(end, mappings, sheet);
            if matches!(a.node, Node::InvalidReference { .. })
                || matches!(b.node, Node::InvalidReference { .. })
            {
                return invalid(e);
            }
            let mut out = e.clone();
            out.node = Node::RangeReference {
                start: Box::new(a),
                end: Box::new(b),
            };
            return out;
        }
        map_children(e, &mut |c| Ok(visit(c, mappings, inherited))).expect("infallible traversal")
    }
    visit(e, mappings, None)
}
pub fn shift_f4(reference: &ParsedCellReference) -> ParsedCellReference {
    let mut r = reference.clone();
    (r.absolute_row, r.absolute_column) = match (r.absolute_row, r.absolute_column) {
        (false, false) => (true, true),
        (true, true) => (true, false),
        (true, false) => (false, true),
        (false, true) => (false, false),
    };
    r
}
#[cfg(test)]
mod tests {
    use super::*;
    fn parse(s: &str) -> Expr {
        parse_source(s).unwrap()
    }
    #[test]
    fn source_roundtrip_preserves_precedence_anchors_and_groups() {
        for input in [
            "1+2*3",
            "1&(2+3)",
            "(1+2)*3",
            "2^3^2",
            "-2^2",
            "2^-2",
            "50%+2",
            "SUM($A$1:B2,,3)",
            "LAMBDA(x,x+1)(3)",
            "LET(x,3,x^2)",
            "{1,2;3,4}",
            "SUM((A1,B2))",
            "A1:C3 B2:D4",
            "Table1[[#Headers],[销售额]]",
            "Table1[@[销售额]]",
        ] {
            let a = parse(input);
            let formatted = format_editor(&a);
            let b = parse(&formatted);
            assert_eq!(
                crate::parser::parse(
                    input,
                    &CellAddress {
                        sheet_id: "sheet".into(),
                        row: 0,
                        column: 0
                    }
                )
                .map(|e| format!("{e:?}")),
                crate::parser::parse(
                    &formatted,
                    &CellAddress {
                        sheet_id: "sheet".into(),
                        row: 0,
                        column: 0
                    }
                )
                .map(|e| format!("{e:?}")),
                "{input} => {formatted}"
            );
            assert_eq!(format_editor(&b), formatted, "format idempotence: {input}");
        }
    }
    #[test]
    fn error_array_and_missing_argument_dto_are_complete() {
        let e = parse("=IF(TRUE,{1,2;3,4},#DIV/0!)");
        let j = serde_json::to_value(&e).unwrap();
        assert_eq!(j["arguments"][1]["type"], "array-literal");
        assert_eq!(j["arguments"][2]["code"], "#DIV/0!");
        let e = parse("=IF(TRUE,,)");
        let Node::FunctionCall { arguments, .. } = e.node else {
            panic!()
        };
        assert_eq!(arguments.len(), 3);
        assert!(matches!(arguments[1].node, Node::MissingArgument {}));
    }
    #[test]
    fn spans_use_utf16_and_preserve_reference_anchors() {
        let e = parse("=\"😀\"&'数据😀'!$B$2:C3");
        let Node::BinaryExpression { right, .. } = e.node else {
            panic!()
        };
        assert_eq!(right.span.start, 6);
        assert_eq!(right.span.end, 20);
        let Node::RangeReference { start, end } = right.node else {
            panic!()
        };
        assert_eq!(start.span.start, 6);
        assert_eq!(start.span.end, 17);
        assert_eq!(end.span.start, 18);
        let j = serde_json::to_value(start).unwrap();
        assert_eq!(j["reference"]["sheetId"], "数据😀");
        assert_eq!(j["reference"]["absoluteRow"], true);
        assert_eq!(j["reference"]["absoluteColumn"], true);
    }
    #[test]
    fn offsets_and_f4_use_four_states_and_invalidate_bounds() {
        assert_eq!(
            format_editor(&offset_editor(&parse("$A1+B$2+$C$3"), 2, 3)),
            "$A3+E$2+$C$3"
        );
        assert_eq!(
            format_editor(&offset_editor(&parse("A1:B2"), -1, 0)),
            "#REF!"
        );
        assert_eq!(
            format_editor(&offset_editor(&parse("$A:$B"), 2, 1)),
            "$A:$B"
        );
        let (original, _) = endpoint("A1", None).unwrap();
        let mut r = original.clone();
        for expected in [(true, true), (true, false), (false, true), (false, false)] {
            r = shift_f4(&r);
            assert_eq!((r.absolute_row, r.absolute_column), expected);
        }
        assert_eq!(r, original);
    }
    #[test]
    fn mapping_range_endpoint_invalidates_range_and_uses_inherited_sheet() {
        let e = parse("'数据表'!A1:B2");
        let refs = references_editor(&e);
        assert_eq!(refs.len(), 2);
        assert_eq!(refs[1].sheet_id.as_deref(), Some("数据表"));
        assert_eq!(
            format_editor(&remap_editor(&e, &[(refs[1].clone(), None)])),
            "#REF!"
        );
    }
    #[test]
    fn malformed_syntax_and_budgets_fail_closed() {
        for source in [
            "Sheet1:Sheet3!A1",
            "'Sheet 1:Sheet 3'!A1",
            "'[Book.xlsx]Sheet1'!A1",
            "[Book.xlsx]Sheet1!A1",
        ] {
            assert_eq!(
                parse_source(source).unwrap_err().code,
                "UNSUPPORTED_FEATURE"
            );
        }
        for source in [
            "=",
            "==1",
            "A0",
            "$A$$1",
            "\"unterminated",
            "'Sheet'",
            "SUM(1",
            "{1,2;3}",
            "A1:2",
            "1e",
            "XFE1",
        ] {
            assert!(parse_source(source).is_err(), "must reject {source}");
        }
        assert!(parse_source(&format!("{}1{}", "(".repeat(110), ")".repeat(110))).is_err());
        assert!(parse_source(&"1+".repeat(1024)).is_err());
    }
}
/// Validate an editor DTO before formatting or transforming host-supplied syntax.
/// Parsing calls this same boundary, so source and serialized ASTs share limits.
pub fn validate_editor(expr: &Expr) -> KernelResult<()> {
    let mut pending = vec![(expr, 0usize)];
    let mut count = 0usize;
    while let Some((e, depth)) = pending.pop() {
        count += 1;
        if count > 2048 || depth > 192 {
            return Err(fail("Formula AST budget exceeded"));
        }
        if e.span.start > e.span.end {
            return Err(fail("Invalid formula source span"));
        }
        match &e.node {
            Node::NumberLiteral { value } if !value.is_finite() => {
                return Err(fail("Number literal must be finite"));
            }
            Node::CellReference { reference } => {
                if reference.row >= MAX_ROWS
                    || reference.column >= MAX_COLUMNS
                    || reference
                        .sheet_id
                        .as_ref()
                        .is_some_and(|s| s.trim().is_empty())
                {
                    return Err(fail("Invalid cell reference DTO"));
                }
            }
            Node::RangeReference { start, end } => {
                if !matches!(start.node, Node::CellReference { .. })
                    || !matches!(end.node, Node::CellReference { .. })
                {
                    return Err(fail("Range endpoints must be cell references"));
                }
            }
            Node::WholeColumnReference {
                start_column,
                end_column,
                sheet_id,
                ..
            } => {
                if *start_column >= MAX_COLUMNS
                    || *end_column >= MAX_COLUMNS
                    || sheet_id.as_ref().is_some_and(|s| s.trim().is_empty())
                {
                    return Err(fail("Invalid whole-column reference DTO"));
                }
            }
            Node::WholeRowReference {
                start_row,
                end_row,
                sheet_id,
                ..
            } => {
                if *start_row >= MAX_ROWS
                    || *end_row >= MAX_ROWS
                    || sheet_id.as_ref().is_some_and(|s| s.trim().is_empty())
                {
                    return Err(fail("Invalid whole-row reference DTO"));
                }
            }
            Node::UnaryExpression { operator, .. }
                if !matches!(operator.as_str(), "+" | "-" | "%" | "@") =>
            {
                return Err(fail("Unknown unary operator"));
            }
            Node::BinaryExpression { operator, .. }
                if !matches!(
                    operator.as_str(),
                    "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | "<=" | ">" | ">="
                ) =>
            {
                return Err(fail("Unknown binary operator"));
            }
            Node::ArrayLiteral { rows } => {
                if rows.is_empty()
                    || rows[0].is_empty()
                    || rows.iter().any(|r| r.len() != rows[0].len())
                {
                    return Err(fail("Array must be nonempty and rectangular"));
                }
            }
            Node::InvalidReference { code } if code != "#REF!" => {
                return Err(fail("Invalid-reference code must be #REF!"));
            }
            Node::ErrorLiteral { code }
                if !matches!(
                    code.as_str(),
                    "#GETTING_DATA"
                        | "#BLOCKED!"
                        | "#CONNECT!"
                        | "#UNKNOWN!"
                        | "#PYTHON!"
                        | "#SPILL!"
                        | "#VALUE!"
                        | "#DIV/0!"
                        | "#FIELD!"
                        | "#CALC!"
                        | "#NULL!"
                        | "#NAME?"
                        | "#BUSY!"
                        | "#NUM!"
                        | "#N/A"
                ) =>
            {
                return Err(fail("Unknown formula error literal"));
            }
            Node::FunctionCall { name, .. } | Node::NameReference { name } if name.is_empty() => {
                return Err(fail("Formula identifier cannot be empty"));
            }
            Node::TableReference {
                table_name,
                specifier,
                column_name,
                column_end_name,
                ..
            } => {
                if table_name.is_empty()
                    || specifier.as_ref().is_some_and(|s| {
                        !matches!(s.as_str(), "all" | "headers" | "data" | "totals")
                    })
                    || column_name.as_ref().is_some_and(|s| s.is_empty())
                    || column_end_name.is_some() && column_name.is_none()
                {
                    return Err(fail("Invalid structured reference DTO"));
                }
            }
            Node::ReferenceUnion { references } => {
                if references.len() < 2 || references.iter().any(|r| !is_ref(r)) {
                    return Err(fail("Union requires at least two references"));
                }
            }
            Node::ReferenceIntersection { left, right } => {
                if !is_ref(left) || !is_ref(right) {
                    return Err(fail("Intersection requires reference operands"));
                }
            }
            _ => {}
        }
        pending.extend(children(e).into_iter().map(|child| (child, depth + 1)));
    }
    Ok(())
}
