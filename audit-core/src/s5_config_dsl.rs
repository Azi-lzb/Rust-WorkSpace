//! S5 configuration DSL parser and static validation.
//!
//! This module is a faithful Rust port of the parse/validation half of the
//! source project's `s1audit_configured.dsl` (the `evaluate` runtime belongs
//! to the later rule-engine slice). It parses audit rule expressions and value
//! dictionaries, then applies the same static checks: allowed calls, call
//! arity/keywords, static type checking, and dictionary dependency checks with
//! cycle detection. Error messages mirror the Python implementation verbatim
//! so real-configuration parity can compare them exactly.
//!
//! Known Unicode edge delta: Python's `str.isdigit` accepts every Unicode
//! decimal digit (for example Arabic-Indic digits, which `Decimal` then
//! accepts too), while this port accepts ASCII digits plus the superscript
//! trio (which then fail number validation with the same "无效数字" error the
//! Python `Decimal` constructor raises for those characters).

use std::collections::BTreeSet;

/// Which Python exception class a DSL failure corresponds to.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum DslErrorKind {
    /// `DSLParseError`
    Parse,
    /// `ConfigurationError`
    Configuration,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DslError {
    pub kind: DslErrorKind,
    pub message: String,
}

impl DslError {
    pub fn parse(message: impl Into<String>) -> Self {
        Self {
            kind: DslErrorKind::Parse,
            message: message.into(),
        }
    }

    pub fn configuration(message: impl Into<String>) -> Self {
        Self {
            kind: DslErrorKind::Configuration,
            message: message.into(),
        }
    }
}

impl std::fmt::Display for DslError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "{}", self.message)
    }
}

impl std::error::Error for DslError {}

#[derive(Clone, Debug, PartialEq)]
pub struct LiteralNode {
    pub value: LiteralValue,
    /// `"number"`, `"string"`, or `"boolean"`, mirroring the Python `kind`.
    pub kind: &'static str,
}

#[derive(Clone, Debug, PartialEq)]
pub enum LiteralValue {
    /// Raw numeric text as written; numeric evaluation belongs to the later
    /// rule-engine slice.
    Number(String),
    String(String),
    Boolean(bool),
}

#[derive(Clone, Debug, PartialEq)]
pub struct VariableNode {
    pub name: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct TupleNode {
    pub values: Vec<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum UnaryOperator {
    Not,
    Plus,
    Minus,
}

impl UnaryOperator {
    fn as_str(self) -> &'static str {
        match self {
            Self::Not => "NOT",
            Self::Plus => "+",
            Self::Minus => "-",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum BinaryOperator {
    And,
    Or,
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
    Add,
    Subtract,
    Multiply,
    Divide,
}

impl BinaryOperator {
    fn as_str(self) -> &'static str {
        match self {
            Self::And => "AND",
            Self::Or => "OR",
            Self::Equal => "=",
            Self::NotEqual => "<>",
            Self::Less => "<",
            Self::LessEqual => "<=",
            Self::Greater => ">",
            Self::GreaterEqual => ">=",
            Self::Add => "+",
            Self::Subtract => "-",
            Self::Multiply => "*",
            Self::Divide => "/",
        }
    }

    fn is_comparison(self) -> bool {
        matches!(
            self,
            Self::Equal
                | Self::NotEqual
                | Self::Less
                | Self::LessEqual
                | Self::Greater
                | Self::GreaterEqual
        )
    }

    fn from_comparison_text(text: &str) -> Option<Self> {
        match text {
            "=" | "==" => Some(Self::Equal),
            "<>" | "!=" => Some(Self::NotEqual),
            "<" => Some(Self::Less),
            "<=" => Some(Self::LessEqual),
            ">" => Some(Self::Greater),
            ">=" => Some(Self::GreaterEqual),
            _ => None,
        }
    }

    fn from_arithmetic_text(text: &str) -> Option<Self> {
        match text {
            "+" => Some(Self::Add),
            "-" => Some(Self::Subtract),
            "*" => Some(Self::Multiply),
            "/" => Some(Self::Divide),
            _ => None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct CallNode {
    pub name: String,
    pub arguments: Vec<Node>,
    /// Keyword arguments in source order, mirroring the Python tuple pairs.
    pub keywords: Vec<(String, String)>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum Node {
    Literal(LiteralNode),
    Variable(VariableNode),
    Tuple(TupleNode),
    Unary {
        operator: UnaryOperator,
        operand: Box<Node>,
    },
    Binary {
        operator: BinaryOperator,
        left: Box<Node>,
        right: Box<Node>,
    },
    Call(CallNode),
}

const ALLOWED_CALLS: [&str; 14] = [
    "V_CUR", "V_PRE", "V_EXT", "ABS", "ROUND", "MIN", "MAX", "IF", "AND", "OR", "NOT", "SUM",
    "ISBLANK", "ISNUMBER",
];

const ALLOWED_KEYWORDS: [&str; 6] = [
    "source",
    "sheet",
    "unit",
    "source_unit",
    "standard_unit",
    "metric_type",
];

const KEYWORD_CALLS: [&str; 3] = ["V_CUR", "V_PRE", "V_EXT"];

fn call_arity(name: &str) -> Option<(usize, usize)> {
    Some(match name {
        "V_CUR" | "V_PRE" => (2, 3),
        "V_EXT" | "ABS" | "NOT" | "ISBLANK" | "ISNUMBER" => (1, 1),
        "ROUND" => (2, 2),
        "IF" => (3, 3),
        "MIN" | "MAX" | "AND" | "OR" | "SUM" => (1, 99),
        _ => return None,
    })
}

fn is_identifier_start(character: char) -> bool {
    character == '_' || character.is_alphabetic() || ('\u{3400}'..='\u{9fff}').contains(&character)
}

fn is_identifier_char(character: char) -> bool {
    is_identifier_start(character) || character.is_ascii_digit()
}

/// Python `str.isdigit` on ASCII plus the superscript trio.
fn is_digit_like(character: char) -> bool {
    character.is_ascii_digit() || matches!(character, '²' | '³' | '¹')
}

#[derive(Clone, Debug)]
struct Token {
    kind: &'static str,
    value: String,
    position: usize,
}

fn tokenize(text: &str) -> Result<Vec<Token>, DslError> {
    let characters: Vec<char> = text.chars().collect();
    let mut tokens = Vec::new();
    let mut index = 0;
    while index < characters.len() {
        let character = characters[index];
        if character.is_whitespace() {
            index += 1;
            continue;
        }
        if character == '"' {
            let start = index;
            index += 1;
            let mut parts = Vec::new();
            let mut closed = false;
            while index < characters.len() {
                let current = characters[index];
                if current == '"' {
                    index += 1;
                    closed = true;
                    break;
                }
                if current == '\\' {
                    index += 1;
                    let replacement = match characters.get(index) {
                        Some('"') => Some('"'),
                        Some('\\') => Some('\\'),
                        Some('n') => Some('\n'),
                        Some('r') => Some('\r'),
                        Some('t') => Some('\t'),
                        _ => None,
                    };
                    let Some(replacement) = replacement else {
                        return Err(DslError::parse(format!(
                            "字符串包含不支持的转义，位置 {}",
                            index
                        )));
                    };
                    parts.push(replacement);
                    index += 1;
                    continue;
                }
                parts.push(current);
                index += 1;
            }
            if !closed {
                return Err(DslError::parse(format!(
                    "字符串缺少结束引号，位置 {}",
                    start
                )));
            }
            tokens.push(Token {
                kind: "STRING",
                value: parts.into_iter().collect(),
                position: start,
            });
            continue;
        }
        let starts_number = is_digit_like(character)
            || characters
                .get(index + 1)
                .is_some_and(|next| character == '.' && is_digit_like(*next));
        if starts_number {
            let start = index;
            let mut saw_dot = character == '.';
            index += 1;
            while index < characters.len() {
                let current = characters[index];
                if is_digit_like(current) {
                    index += 1;
                } else if current == '.' && !saw_dot {
                    saw_dot = true;
                    index += 1;
                } else {
                    break;
                }
            }
            tokens.push(Token {
                kind: "NUMBER",
                value: characters[start..index].iter().collect(),
                position: start,
            });
            continue;
        }
        if is_identifier_start(character) {
            let start = index;
            index += 1;
            while index < characters.len() && is_identifier_char(characters[index]) {
                index += 1;
            }
            let value: String = characters[start..index].iter().collect();
            let keyword = value.to_ascii_uppercase();
            let kind = match keyword.as_str() {
                "AND" | "OR" | "NOT" | "TRUE" | "FALSE" => match keyword.as_str() {
                    "AND" => "AND",
                    "OR" => "OR",
                    "NOT" => "NOT",
                    "TRUE" => "TRUE",
                    _ => "FALSE",
                },
                _ => "IDENT",
            };
            tokens.push(Token {
                kind,
                value,
                position: start,
            });
            continue;
        }
        if index + 1 < characters.len() && matches!(character, '=' | '<' | '>' | '!') {
            let pair: String = characters[index..index + 2].iter().collect();
            if matches!(pair.as_str(), "==" | "<>" | "!=" | "<=" | ">=") {
                tokens.push(Token {
                    kind: "OP",
                    value: pair,
                    position: index,
                });
                index += 2;
                continue;
            }
        }
        if matches!(
            character,
            '+' | '-' | '*' | '/' | '=' | '<' | '>' | '!' | '(' | ')' | '{' | '}' | ',' | ':'
        ) {
            let kind = match character {
                '+' | '-' | '*' | '/' | '=' | '<' | '>' | '!' => "OP",
                '(' => "LPAREN",
                ')' => "RPAREN",
                '{' => "LBRACE",
                '}' => "RBRACE",
                ',' => "COMMA",
                ':' => "COLON",
                _ => unreachable!(),
            };
            tokens.push(Token {
                kind,
                value: character.to_string(),
                position: index,
            });
            index += 1;
            continue;
        }
        return Err(DslError::parse(format!(
            "不支持的字符 {:?}，位置 {}",
            character, index
        )));
    }
    tokens.push(Token {
        kind: "EOF",
        value: String::new(),
        position: characters.len(),
    });
    Ok(tokens)
}

struct Parser {
    tokens: Vec<Token>,
    index: usize,
}

impl Parser {
    fn new(text: &str) -> Result<Self, DslError> {
        Ok(Self {
            tokens: tokenize(text)?,
            index: 0,
        })
    }

    fn current(&self) -> &Token {
        &self.tokens[self.index]
    }

    fn peek(&self, distance: usize) -> &Token {
        &self.tokens[(self.index + distance).min(self.tokens.len() - 1)]
    }

    fn advance(&mut self) -> Token {
        let token = self.tokens[self.index].clone();
        self.index += 1;
        token
    }

    fn accept(&mut self, kind: &str, value: Option<&str>) -> bool {
        let token = self.current();
        if token.kind == kind && (value.is_none() || token.value == value.unwrap_or_default()) {
            self.index += 1;
            return true;
        }
        false
    }

    fn expect(&mut self, kind: &str, value: Option<&str>) -> Result<Token, DslError> {
        if self.accept(kind, value) {
            return Ok(self.tokens[self.index - 1].clone());
        }
        let expected = value.unwrap_or(kind);
        Err(DslError::parse(format!(
            "期望 {}，位置 {}",
            expected,
            self.current().position
        )))
    }

    fn parse_expression(&mut self) -> Result<Node, DslError> {
        let node = self.parse_or()?;
        self.expect("EOF", None)?;
        Ok(node)
    }

    fn parse_dictionary(&mut self) -> Result<Vec<(String, Node)>, DslError> {
        self.expect("LBRACE", None)?;
        let mut result: Vec<(String, Node)> = Vec::new();
        if self.accept("RBRACE", None) {
            self.expect("EOF", None)?;
            return Ok(result);
        }
        loop {
            let key = self.expect("STRING", None)?.value;
            if result.iter().any(|(existing, _)| *existing == key) {
                return Err(DslError::configuration(format!("取值字典重复变量：{key}")));
            }
            self.expect("COLON", None)?;
            let value = self.parse_or()?;
            result.push((key, value));
            if self.accept("COMMA", None) {
                if self.accept("RBRACE", None) {
                    break;
                }
                continue;
            }
            self.expect("RBRACE", None)?;
            break;
        }
        self.expect("EOF", None)?;
        for (_, node) in &result {
            validate_node(node)?;
        }
        Ok(result)
    }

    fn parse_or(&mut self) -> Result<Node, DslError> {
        let mut node = self.parse_and()?;
        while self.accept("OR", None) {
            let right = self.parse_and()?;
            node = Node::Binary {
                operator: BinaryOperator::Or,
                left: Box::new(node),
                right: Box::new(right),
            };
        }
        Ok(node)
    }

    fn parse_and(&mut self) -> Result<Node, DslError> {
        let mut node = self.parse_comparison()?;
        while self.accept("AND", None) {
            let right = self.parse_comparison()?;
            node = Node::Binary {
                operator: BinaryOperator::And,
                left: Box::new(node),
                right: Box::new(right),
            };
        }
        Ok(node)
    }

    fn parse_comparison(&mut self) -> Result<Node, DslError> {
        let mut node = self.parse_additive()?;
        if self.current().kind == "OP" {
            if let Some(operator) = BinaryOperator::from_comparison_text(&self.current().value) {
                self.advance();
                let right = self.parse_additive()?;
                node = Node::Binary {
                    operator,
                    left: Box::new(node),
                    right: Box::new(right),
                };
                if self.current().kind == "OP"
                    && BinaryOperator::from_comparison_text(&self.current().value).is_some()
                {
                    return Err(DslError::parse(format!(
                        "不支持连续比较，位置 {}",
                        self.current().position
                    )));
                }
            }
        }
        Ok(node)
    }

    fn parse_additive(&mut self) -> Result<Node, DslError> {
        let mut node = self.parse_multiplicative()?;
        while self.current().kind == "OP" && matches!(self.current().value.as_str(), "+" | "-") {
            let text = self.advance().value;
            let operator = BinaryOperator::from_arithmetic_text(&text).unwrap();
            let right = self.parse_multiplicative()?;
            node = Node::Binary {
                operator,
                left: Box::new(node),
                right: Box::new(right),
            };
        }
        Ok(node)
    }

    fn parse_multiplicative(&mut self) -> Result<Node, DslError> {
        let mut node = self.parse_unary()?;
        while self.current().kind == "OP" && matches!(self.current().value.as_str(), "*" | "/") {
            let text = self.advance().value;
            let operator = BinaryOperator::from_arithmetic_text(&text).unwrap();
            let right = self.parse_unary()?;
            node = Node::Binary {
                operator,
                left: Box::new(node),
                right: Box::new(right),
            };
        }
        Ok(node)
    }

    fn parse_unary(&mut self) -> Result<Node, DslError> {
        if self.current().kind == "NOT" {
            self.advance();
            let operand = self.parse_unary()?;
            return Ok(Node::Unary {
                operator: UnaryOperator::Not,
                operand: Box::new(operand),
            });
        }
        if self.current().kind == "OP" && matches!(self.current().value.as_str(), "+" | "-") {
            let text = self.advance().value;
            let operator = if text == "+" {
                UnaryOperator::Plus
            } else {
                UnaryOperator::Minus
            };
            let operand = self.parse_unary()?;
            return Ok(Node::Unary {
                operator,
                operand: Box::new(operand),
            });
        }
        self.parse_primary()
    }

    fn parse_primary(&mut self) -> Result<Node, DslError> {
        let token = self.current().clone();
        if token.kind == "NUMBER" {
            self.advance();
            if !is_valid_number_text(&token.value) {
                return Err(DslError::parse(format!(
                    "无效数字，位置 {}",
                    token.position
                )));
            }
            return Ok(Node::Literal(LiteralNode {
                value: LiteralValue::Number(token.value),
                kind: "number",
            }));
        }
        if token.kind == "STRING" {
            self.advance();
            return Ok(Node::Literal(LiteralNode {
                value: LiteralValue::String(token.value),
                kind: "string",
            }));
        }
        if matches!(token.kind, "TRUE" | "FALSE") {
            self.advance();
            return Ok(Node::Literal(LiteralNode {
                value: LiteralValue::Boolean(token.kind == "TRUE"),
                kind: "boolean",
            }));
        }
        if token.kind == "LBRACE" {
            self.advance();
            let name = self.expect("IDENT", None)?.value;
            self.expect("RBRACE", None)?;
            return Ok(Node::Variable(VariableNode { name }));
        }
        if token.kind == "IDENT" {
            self.advance();
            let mut name = token.value.to_ascii_uppercase();
            name = match name.as_str() {
                "VCUR" => "V_CUR".to_owned(),
                "VPRE" => "V_PRE".to_owned(),
                "VEXT" | "EXT" => "V_EXT".to_owned(),
                _ => name,
            };
            if !self.accept("LPAREN", None) {
                return Err(DslError::parse(format!(
                    "标识符必须是受支持的函数调用，位置 {}",
                    token.position
                )));
            }
            let mut arguments = Vec::new();
            let mut keywords: Vec<(String, String)> = Vec::new();
            if !self.accept("RPAREN", None) {
                loop {
                    if self.current().kind == "IDENT"
                        && self.peek(1).kind == "OP"
                        && self.peek(1).value == "="
                    {
                        let key = self.advance().value;
                        self.advance();
                        if self.current().kind != "STRING" {
                            return Err(DslError::parse(format!(
                                "参数 {key}= 的值必须是字符串，位置 {}",
                                self.current().position
                            )));
                        }
                        keywords.push((key, self.advance().value));
                    } else {
                        arguments.push(self.parse_argument()?);
                    }
                    if self.accept("COMMA", None) {
                        if self.accept("RPAREN", None) {
                            break;
                        }
                        continue;
                    }
                    self.expect("RPAREN", None)?;
                    break;
                }
            }
            let call = CallNode {
                name,
                arguments,
                keywords,
            };
            validate_call(&call)?;
            return Ok(Node::Call(call));
        }
        if token.kind == "LPAREN" {
            self.advance();
            let node = self.parse_or()?;
            self.expect("RPAREN", None)?;
            return Ok(node);
        }
        Err(DslError::parse(format!(
            "此处需要值，位置 {}",
            token.position
        )))
    }

    fn parse_argument(&mut self) -> Result<Node, DslError> {
        if self.current().kind == "LPAREN"
            && self.peek(1).kind == "STRING"
            && self.peek(2).kind == "COMMA"
        {
            self.advance();
            let mut values = vec![self.expect("STRING", None)?.value];
            while self.accept("COMMA", None) {
                if self.current().kind == "RPAREN" {
                    break;
                }
                if self.current().kind != "STRING" {
                    return Err(DslError::parse(format!(
                        "路径元组只能包含字符串，位置 {}",
                        self.current().position
                    )));
                }
                values.push(self.advance().value);
            }
            self.expect("RPAREN", None)?;
            if values.iter().any(|value| value.trim().is_empty()) {
                return Err(DslError::parse("路径元组各层不能为空"));
            }
            return Ok(Node::Tuple(TupleNode { values }));
        }
        self.parse_or()
    }
}

/// Python `Decimal` accepts digits with at most one decimal point; the token
/// grammar already guarantees the shape, so this only rejects the superscript
/// tokens Python's `isdigit` also accepts but `Decimal` rejects.
fn is_valid_number_text(text: &str) -> bool {
    let mut dots = 0;
    let mut digits = 0;
    for character in text.chars() {
        if character.is_ascii_digit() {
            digits += 1;
        } else if character == '.' {
            dots += 1;
        } else {
            return false;
        }
    }
    digits >= 1 && dots <= 1
}

pub fn validate_call(call: &CallNode) -> Result<(), DslError> {
    if !ALLOWED_CALLS.contains(&call.name.as_str()) {
        return Err(DslError::configuration(format!(
            "不支持的函数：{}",
            call.name
        )));
    }
    let keyword_names: Vec<&str> = call
        .keywords
        .iter()
        .map(|(name, _)| name.as_str())
        .collect();
    let unique_names = keyword_names.iter().collect::<BTreeSet<_>>().len() == keyword_names.len();
    if !unique_names {
        return Err(DslError::configuration(format!(
            "函数 {} 存在重复具名参数",
            call.name
        )));
    }
    if keyword_names
        .iter()
        .any(|name| !ALLOWED_KEYWORDS.contains(name))
    {
        return Err(DslError::configuration(format!(
            "函数 {} 的具名参数不受支持",
            call.name
        )));
    }
    if !KEYWORD_CALLS.contains(&call.name.as_str()) && !call.keywords.is_empty() {
        return Err(DslError::configuration(format!(
            "函数 {} 不接受具名参数",
            call.name
        )));
    }
    let (minimum, maximum) = call_arity(&call.name).expect("allowed calls always have an arity");
    if call.arguments.len() < minimum || call.arguments.len() > maximum {
        return Err(DslError::configuration(format!(
            "函数 {} 参数个数错误",
            call.name
        )));
    }
    if KEYWORD_CALLS.contains(&call.name.as_str()) {
        let mut path_arguments: &[Node] = &call.arguments;
        if matches!(call.name.as_str(), "V_CUR" | "V_PRE") && call.arguments.len() == 3 {
            let sheet_argument = &call.arguments[0];
            let is_nonempty_string = matches!(
                sheet_argument,
                Node::Literal(LiteralNode {
                    value: LiteralValue::String(value),
                    kind: "string",
                }) if !value.trim().is_empty()
            );
            if !is_nonempty_string {
                return Err(DslError::configuration(format!(
                    "函数 {} 的第一个参数必须是非空工作表名称",
                    call.name
                )));
            }
            if keyword_names.iter().any(|name| *name == "sheet") {
                return Err(DslError::configuration(format!(
                    "函数 {} 已通过第一个参数指定工作表，不能再使用 sheet=",
                    call.name
                )));
            }
            path_arguments = &call.arguments[1..];
        }
        for argument in path_arguments {
            match argument {
                Node::Literal(LiteralNode {
                    value: LiteralValue::String(_),
                    kind: "string",
                }) => {}
                Node::Tuple(_) => {}
                Node::Literal(_) => {
                    return Err(DslError::configuration(format!(
                        "函数 {} 的定位参数必须是字符串",
                        call.name
                    )));
                }
                _ => {
                    return Err(DslError::configuration(format!(
                        "函数 {} 的定位参数必须是字符串常量或字符串元组",
                        call.name
                    )));
                }
            }
        }
        for (_, value) in &call.keywords {
            if value.trim().is_empty() {
                return Err(DslError::configuration("source= 和 sheet= 不能为空"));
            }
        }
    }
    Ok(())
}

pub fn validate_node(node: &Node) -> Result<(), DslError> {
    match node {
        Node::Call(call) => {
            validate_call(call)?;
            for argument in &call.arguments {
                validate_node(argument)?;
            }
            Ok(())
        }
        Node::Unary { operand, .. } => validate_node(operand),
        Node::Binary { left, right, .. } => {
            validate_node(left)?;
            validate_node(right)
        }
        _ => Ok(()),
    }
}

pub fn parse_expression(text: &str) -> Result<Node, DslError> {
    if text.trim().is_empty() {
        return Err(DslError::parse("表达式不能为空"));
    }
    let mut parser = Parser::new(text)?;
    let node = parser.parse_expression()?;
    validate_node(&node)?;
    Ok(node)
}

/// Insertion-ordered variable dictionary parsed from `取值字典` text.
pub fn parse_dictionary(text: Option<&str>) -> Result<Vec<(String, Node)>, DslError> {
    let Some(text) = text else {
        return Ok(Vec::new());
    };
    if text.trim().is_empty() {
        return Ok(Vec::new());
    }
    Parser::new(text)?.parse_dictionary()
}

pub fn variables_in(node: &Node) -> BTreeSet<String> {
    let mut found = BTreeSet::new();
    collect_variables(node, &mut found);
    found
}

fn collect_variables(node: &Node, found: &mut BTreeSet<String>) {
    match node {
        Node::Variable(variable) => {
            found.insert(variable.name.clone());
        }
        Node::Call(call) => {
            for argument in &call.arguments {
                collect_variables(argument, found);
            }
        }
        Node::Unary { operand, .. } => collect_variables(operand, found),
        Node::Binary { left, right, .. } => {
            collect_variables(left, found);
            collect_variables(right, found);
        }
        _ => {}
    }
}

pub fn calls_in(node: &Node) -> Vec<&CallNode> {
    let mut found = Vec::new();
    collect_calls(node, &mut found);
    found
}

fn collect_calls<'node>(node: &'node Node, found: &mut Vec<&'node CallNode>) {
    match node {
        Node::Call(call) => {
            found.push(call);
            for argument in &call.arguments {
                collect_calls(argument, found);
            }
        }
        Node::Unary { operand, .. } => collect_calls(operand, found),
        Node::Binary { left, right, .. } => {
            collect_calls(left, found);
            collect_calls(right, found);
        }
        _ => {}
    }
}

pub fn validate_dictionary_dependencies(dictionary: &[(String, Node)]) -> Result<(), DslError> {
    let dictionary_names: BTreeSet<&str> =
        dictionary.iter().map(|(name, _)| name.as_str()).collect();
    let dependencies: Vec<(String, BTreeSet<String>)> = dictionary
        .iter()
        .map(|(name, node)| (name.clone(), variables_in(node)))
        .collect();
    for (name, names) in &dependencies {
        let undefined: Vec<&str> = names
            .iter()
            .map(String::as_str)
            .filter(|candidate| !dictionary_names.contains(candidate))
            .collect();
        if !undefined.is_empty() {
            return Err(DslError::configuration(format!(
                "变量 {name} 引用了未定义变量：{}",
                undefined.join(", ")
            )));
        }
    }
    let mut visiting: Vec<String> = Vec::new();
    let mut visited: BTreeSet<String> = BTreeSet::new();
    for (name, _) in dictionary {
        if let Some(error) =
            visit_dictionary_dependency(name, &dependencies, &mut visiting, &mut visited)
        {
            return Err(DslError::configuration(error));
        }
    }
    Ok(())
}

fn visit_dictionary_dependency(
    name: &str,
    dependencies: &[(String, BTreeSet<String>)],
    visiting: &mut Vec<String>,
    visited: &mut BTreeSet<String>,
) -> Option<String> {
    if visiting.iter().any(|entry| entry == name) {
        let start = visiting.iter().position(|entry| entry == name).unwrap();
        let mut cycle: Vec<String> = visiting[start..].to_vec();
        cycle.push(name.to_owned());
        return Some(format!("取值字典存在循环引用：{}", cycle.join(" → ")));
    }
    if visited.contains(name) {
        return None;
    }
    visiting.push(name.to_owned());
    let dependencies_for_name = dependencies
        .iter()
        .find(|(entry, _)| entry == name)
        .map(|(_, names)| names.clone())
        .unwrap_or_default();
    for dependency in dependencies_for_name {
        if let Some(error) =
            visit_dictionary_dependency(&dependency, dependencies, visiting, visited)
        {
            return Some(error);
        }
    }
    visiting.pop();
    visited.insert(name.to_owned());
    None
}

pub fn static_kind(node: &Node) -> Option<&'static str> {
    match node {
        Node::Literal(literal) => Some(literal.kind),
        Node::Variable(_) => None,
        Node::Tuple(_) => Some("tuple"),
        Node::Unary { operator, .. } => match operator {
            UnaryOperator::Not => Some("boolean"),
            _ => Some("number"),
        },
        Node::Binary { operator, .. } => {
            if matches!(operator, BinaryOperator::And | BinaryOperator::Or)
                || operator.is_comparison()
            {
                Some("boolean")
            } else {
                Some("number")
            }
        }
        Node::Call(call) => match call.name.as_str() {
            "V_CUR" | "V_PRE" | "V_EXT" => None,
            "MIN" | "MAX" | "ABS" | "ROUND" | "SUM" => Some("number"),
            "ISBLANK" | "ISNUMBER" => Some("boolean"),
            "IF" => {
                let yes = static_kind(call.arguments.get(1)?)?;
                let no = static_kind(call.arguments.get(2)?)?;
                if yes == no {
                    Some(yes)
                } else {
                    None
                }
            }
            "AND" | "OR" | "NOT" => Some("boolean"),
            _ => None,
        },
    }
}

pub fn validate_static_types(node: &Node, allow_string_literal: bool) -> Result<(), DslError> {
    match node {
        Node::Literal(literal) => {
            if literal.kind == "string" && !allow_string_literal {
                return Err(DslError::configuration(
                    "字符串常量只能用于 V 函数参数或 IF 的结果分支",
                ));
            }
            Ok(())
        }
        Node::Variable(_) | Node::Tuple(_) => Ok(()),
        Node::Unary { operator, operand } => {
            validate_static_types(operand, allow_string_literal)?;
            let expected = match operator {
                UnaryOperator::Not => "boolean",
                _ => "number",
            };
            if let Some(actual) = static_kind(operand) {
                if actual != expected {
                    return Err(DslError::configuration(format!(
                        "{} 运算要求 {}，实际是 {}",
                        operator.as_str(),
                        expected,
                        actual
                    )));
                }
            }
            Ok(())
        }
        Node::Binary {
            operator,
            left,
            right,
        } => {
            validate_static_types(left, allow_string_literal)?;
            validate_static_types(right, allow_string_literal)?;
            let left_kind = static_kind(left);
            let right_kind = static_kind(right);
            if matches!(operator, BinaryOperator::And | BinaryOperator::Or) {
                for kind in [left_kind, right_kind].into_iter().flatten() {
                    if kind != "boolean" {
                        return Err(DslError::configuration(format!(
                            "{} 运算只接受布尔值",
                            operator.as_str()
                        )));
                    }
                }
            } else if matches!(
                operator,
                BinaryOperator::Add
                    | BinaryOperator::Subtract
                    | BinaryOperator::Multiply
                    | BinaryOperator::Divide
            ) {
                for kind in [left_kind, right_kind].into_iter().flatten() {
                    if kind != "number" {
                        return Err(DslError::configuration(format!(
                            "{} 运算只接受数值",
                            operator.as_str()
                        )));
                    }
                }
            } else {
                if let (Some(left_kind), Some(right_kind)) = (left_kind, right_kind) {
                    if left_kind != right_kind {
                        return Err(DslError::configuration(format!(
                            "比较两侧类型不一致：{} / {}",
                            left_kind, right_kind
                        )));
                    }
                }
                if matches!(
                    operator,
                    BinaryOperator::Less
                        | BinaryOperator::LessEqual
                        | BinaryOperator::Greater
                        | BinaryOperator::GreaterEqual
                ) && left_kind.is_some_and(|kind| kind != "number")
                {
                    return Err(DslError::configuration("字符串和布尔值只支持相等比较"));
                }
            }
            Ok(())
        }
        Node::Call(call) => {
            validate_call(call)?;
            match call.name.as_str() {
                "V_CUR" | "V_PRE" | "V_EXT" => Ok(()),
                "IF" => {
                    let condition = &call.arguments[0];
                    validate_static_types(condition, allow_string_literal)?;
                    if let Some(condition_kind) = static_kind(condition) {
                        if condition_kind != "boolean" {
                            return Err(DslError::configuration("IF 的第一个参数必须是布尔值"));
                        }
                    }
                    validate_static_types(&call.arguments[1], true)?;
                    validate_static_types(&call.arguments[2], true)?;
                    let yes_kind = static_kind(&call.arguments[1]);
                    let no_kind = static_kind(&call.arguments[2]);
                    if let (Some(yes_kind), Some(no_kind)) = (yes_kind, no_kind) {
                        if yes_kind != no_kind {
                            return Err(DslError::configuration(format!(
                                "IF 的两个结果分支类型不一致：{} / {}",
                                yes_kind, no_kind
                            )));
                        }
                    }
                    Ok(())
                }
                "ISBLANK" | "ISNUMBER" => validate_static_types(&call.arguments[0], true),
                name => {
                    let expected = if matches!(name, "AND" | "OR" | "NOT") {
                        "boolean"
                    } else {
                        "number"
                    };
                    for argument in &call.arguments {
                        validate_static_types(argument, allow_string_literal)?;
                        if let Some(actual) = static_kind(argument) {
                            if actual != expected {
                                return Err(DslError::configuration(format!(
                                    "{name} 参数要求 {expected}，实际是 {actual}"
                                )));
                            }
                        }
                    }
                    Ok(())
                }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse_ok(text: &str) -> Node {
        parse_expression(text).expect("expression should parse")
    }

    fn parse_err(text: &str) -> DslError {
        parse_expression(text).expect_err("expression should fail")
    }

    #[test]
    fn parses_boolean_expression_with_v_call_and_placeholders() {
        let node = parse_ok("(NOT(ISBLANK({本期值})) AND ({本期值} <> 0))");
        // Prefix NOT is a unary operator in this grammar, so only ISBLANK is
        // a call node, exactly as in the Python AST.
        let calls = calls_in(&node);
        assert_eq!(calls.len(), 1);
        assert_eq!(calls[0].name, "ISBLANK");
        assert!(matches!(static_kind(&node), Some("boolean")));
    }

    #[test]
    fn normalizes_v_function_aliases() {
        let node = parse_ok("VCUR(\"表\", (\"资产@行\", \"负债@行\"))");
        let Node::Call(call) = &node else {
            panic!("expected call node");
        };
        assert_eq!(call.name, "V_CUR");
        assert!(
            matches!(&call.arguments[1], Node::Tuple(TupleNode { values }) if values.len() == 2)
        );
    }

    #[test]
    fn rejects_unknown_function_and_bare_identifiers() {
        let error = parse_err("FOO(1)");
        assert_eq!(error.kind, DslErrorKind::Configuration);
        assert_eq!(error.message, "不支持的函数：FOO");
        let error = parse_err("x");
        assert_eq!(error.kind, DslErrorKind::Parse);
        assert_eq!(error.message, "标识符必须是受支持的函数调用，位置 0");
    }

    #[test]
    fn rejects_chained_comparisons_with_position() {
        let error = parse_err("1 < 2 < 3");
        assert_eq!(error.message, "不支持连续比较，位置 6");
    }

    #[test]
    fn rejects_mismatched_if_branch_kinds_and_string_literals() {
        let node = parse_ok("\"文本\"");
        let error = validate_static_types(&node, false).expect_err("bare string literal");
        assert_eq!(error.kind, DslErrorKind::Configuration);
        assert_eq!(
            error.message,
            "字符串常量只能用于 V 函数参数或 IF 的结果分支"
        );
        parse_ok("IF(TRUE, \"错\", \"对\")");
    }

    #[test]
    fn static_type_checks_match_loader_level_validation() {
        // parse_expression alone accepts these; the configuration loader's
        // validate_static_types pass rejects them, mirroring Python.
        let node = parse_ok("\"文本\"");
        let error = validate_static_types(&node, false).expect_err("bare string literal");
        assert_eq!(error.kind, DslErrorKind::Configuration);
        assert_eq!(
            error.message,
            "字符串常量只能用于 V 函数参数或 IF 的结果分支"
        );
        parse_ok("IF(TRUE, \"错\", \"对\")");
        let node = parse_ok("IF(1 > 0, 1, \"错\")");
        let error = validate_static_types(&node, false).expect_err("mismatched branches");
        assert_eq!(
            error.message,
            "IF 的两个结果分支类型不一致：number / string"
        );
        let node = parse_ok("1 + TRUE");
        let error = validate_static_types(&node, false).expect_err("boolean arithmetic");
        assert_eq!(error.message, "+ 运算只接受数值");
        let node = parse_ok("NOT(1)");
        let error = validate_static_types(&node, false).expect_err("NOT on number");
        assert_eq!(error.message, "NOT 运算要求 boolean，实际是 number");
        let node = parse_ok("\"a\" < \"b\"");
        // The bare string literal on the left is rejected first by the
        // recursive literal check, exactly as in Python.
        let error = validate_static_types(&node, false).expect_err("string ordering");
        assert_eq!(
            error.message,
            "字符串常量只能用于 V 函数参数或 IF 的结果分支"
        );
        let node = parse_ok("IF(TRUE, \"a\", \"a\") < IF(TRUE, \"b\", \"b\")");
        let error = validate_static_types(&node, false).expect_err("string ordering");
        assert_eq!(error.message, "字符串和布尔值只支持相等比较");
        let node = parse_ok("1 < TRUE");
        let error = validate_static_types(&node, false).expect_err("mixed comparison");
        assert_eq!(error.message, "比较两侧类型不一致：number / boolean");
    }

    #[test]
    fn tokenizes_strings_with_escapes_and_reports_positions() {
        parse_ok("\"行\\n文本\"");
        let error = parse_err("\"未闭合");
        assert_eq!(error.message, "字符串缺少结束引号，位置 0");
        let error = parse_err("\"坏\\x\"");
        assert_eq!(error.message, "字符串包含不支持的转义，位置 3");
        let error = parse_err("1 ~ 2");
        assert_eq!(error.message, "不支持的字符 '~'，位置 2");
    }

    #[test]
    fn parses_and_validates_value_dictionaries() {
        let dictionary = parse_dictionary(Some(
            "{\"本期值\": V_CUR(\"表\", \"行@路径\"), \"上期值\": V_PRE(\"表\", \"行@路径\")}",
        ))
        .expect("dictionary should parse");
        assert_eq!(dictionary.len(), 2);
        validate_dictionary_dependencies(&dictionary).expect("dependencies should hold");

        let duplicate = parse_dictionary(Some("{\"a\": 1, \"a\": 2}")).expect_err("duplicate key");
        assert_eq!(duplicate.kind, DslErrorKind::Configuration);
        assert_eq!(duplicate.message, "取值字典重复变量：a");

        let empty = parse_dictionary(None).expect("empty dictionary");
        assert!(empty.is_empty());
    }

    #[test]
    fn rejects_dictionary_cycles_and_undefined_variables() {
        let cyclic = parse_dictionary(Some("{\"a\": {b}, \"b\": {a}}")).expect("parse");
        let error = validate_dictionary_dependencies(&cyclic).expect_err("cycle");
        assert!(
            error.message.starts_with("取值字典存在循环引用："),
            "{error}"
        );

        let undefined = parse_dictionary(Some("{\"a\": {b} + 1}")).expect("parse");
        let error = validate_dictionary_dependencies(&undefined).expect_err("undefined");
        assert_eq!(error.message, "变量 a 引用了未定义变量：b");
    }

    #[test]
    fn keyword_arguments_and_arity_follow_python_semantics() {
        parse_ok("V_EXT(\"指标\", source=\"外部\", unit=\"元\")");
        let error = parse_err("V_CUR(\"表\", \"路径\", \"层\", sheet=\"其他\")");
        assert_eq!(
            error.message,
            "函数 V_CUR 已通过第一个参数指定工作表，不能再使用 sheet="
        );
        let error = parse_err("ABS(1, 2)");
        assert_eq!(error.message, "函数 ABS 参数个数错误");
        let error = parse_err("SUM(1, source=\"x\")");
        assert_eq!(error.message, "函数 SUM 不接受具名参数");
    }

    #[test]
    fn numeric_literals_reject_superscripts_like_python_decimal() {
        let node = parse_ok("1.5 + .5 + 1.");
        assert!(matches!(
            &node,
            Node::Binary {
                operator: BinaryOperator::Add,
                ..
            }
        ));
        let error = parse_err("²");
        assert_eq!(error.message, "无效数字，位置 0");
    }
}
