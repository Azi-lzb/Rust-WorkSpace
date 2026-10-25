//! S5 audit-expression evaluation runtime.
//!
//! Faithful Rust port of the evaluation half of
//! `s1audit_configured.dsl` (`evaluate` and its helpers) on top of the parse
//! tree from [`crate::s5_config_dsl`]. Two numeric modes follow the Python
//! implementation:
//!
//! - `excel_binary64`: f64 arithmetic converged to Excel's 15 significant
//!   digits through `excel_round15`;
//! - `decimal`: a 28-significant-digit decimal context with banker's rounding
//!   for `+`, `-`, `*`, `/` and ROUND_HALF_UP for `ROUND` quantization,
//!   mirroring Python's default `decimal` context.
//!
//! The decimal mode is backed by an internal digit-vector decimal. Business
//! amounts stay far below the working range; values that overflow it fail with
//! an explicit error instead of silently degrading.

use crate::s5_config_dsl::Node;
use serde_json::{json, Value as JsonValue};
use std::collections::BTreeMap;

pub const CONTEXT_DIGITS: usize = 28;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum EvalErrorKind {
    /// `EvaluationError`
    Evaluation,
    /// `ConfigurationError` (non-constant V arguments, unsupported calls)
    Configuration,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct EvalError {
    pub kind: EvalErrorKind,
    pub message: String,
}

impl EvalError {
    fn evaluation(message: impl Into<String>) -> Self {
        Self {
            kind: EvalErrorKind::Evaluation,
            message: message.into(),
        }
    }

    fn configuration(message: impl Into<String>) -> Self {
        Self {
            kind: EvalErrorKind::Configuration,
            message: message.into(),
        }
    }
}

impl std::fmt::Display for EvalError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "{}", self.message)
    }
}

impl std::error::Error for EvalError {}

/// Excel's 15-significant-digit convergence (`%.15g` round-trip).
pub fn excel_round15(value: f64) -> f64 {
    if value == 0.0 || !value.is_finite() {
        return value;
    }
    format!("{value:.14e}").parse::<f64>().unwrap_or(value)
}

/// Plain-notation decimal: value = signed digit vector × 10^exponent.
#[derive(Clone, Debug)]
pub struct Decimal28 {
    negative: bool,
    /// Digits most-significant first without leading zeros; empty means zero.
    digits: Vec<u8>,
    exponent: i32,
}

impl PartialEq for Decimal28 {
    fn eq(&self, other: &Self) -> bool {
        self.cmp_value(other) == std::cmp::Ordering::Equal
    }
}

impl Decimal28 {
    pub fn zero() -> Self {
        Self {
            negative: false,
            digits: Vec::new(),
            exponent: 0,
        }
    }

    pub fn from_i64(value: i64) -> Self {
        let negative = value < 0;
        let mut digits: Vec<u8> = value
            .unsigned_abs()
            .to_string()
            .bytes()
            .map(|byte| byte - b'0')
            .collect();
        strip_leading_zeros(&mut digits);
        Self {
            negative: negative && !digits.is_empty(),
            digits,
            exponent: 0,
        }
    }

    pub fn parse(text: &str) -> Result<Self, String> {
        let text = text.trim();
        let (negative, rest) = match text.strip_prefix('-') {
            Some(rest) => (true, rest),
            None => (false, text.strip_prefix('+').unwrap_or(text)),
        };
        let mut digits: Vec<u8> = Vec::new();
        let mut exponent: i32 = 0;
        let mut seen_dot = false;
        for character in rest.chars() {
            match character {
                '0'..='9' => {
                    digits.push(character as u8 - b'0');
                    if seen_dot {
                        exponent -= 1;
                    }
                }
                '.' if !seen_dot => seen_dot = true,
                _ => return Err(format!("无效数字：{text}")),
            }
        }
        strip_leading_zeros(&mut digits);
        Ok(Self {
            negative: negative && !digits.is_empty(),
            digits,
            exponent,
        })
    }

    pub fn is_zero(&self) -> bool {
        self.digits.is_empty()
    }

    fn magnitude(&self) -> Vec<u8> {
        self.digits.clone()
    }

    /// Python `str(Decimal)` for the plain-notation values reachable here.
    pub fn to_canonical_string(&self) -> String {
        if self.digits.is_empty() {
            return "0".to_owned();
        }
        let mut text: String = self
            .digits
            .iter()
            .map(|digit| (b'0' + digit) as char)
            .collect();
        if self.exponent > 0 {
            text.push_str(&"0".repeat(self.exponent as usize));
        } else if self.exponent < 0 {
            let point = text.len() as i32 + self.exponent;
            if point <= 0 {
                text = format!("0.{}{text}", "0".repeat((-point) as usize));
            } else {
                text.insert(point as usize, '.');
            }
        }
        if self.negative {
            format!("-{text}")
        } else {
            text
        }
    }

    pub fn to_f64(&self) -> f64 {
        self.to_canonical_string().parse::<f64>().unwrap_or(0.0)
    }

    fn rescale_to(&self, exponent: i32) -> Vec<u8> {
        let shift = (self.exponent - exponent).max(0) as usize;
        let mut digits = self.digits.clone();
        digits.extend(std::iter::repeat(0).take(shift));
        strip_leading_zeros(&mut digits);
        digits
    }

    fn magnitude_compare(left: &[u8], right: &[u8]) -> std::cmp::Ordering {
        use std::cmp::Ordering;
        match left.len().cmp(&right.len()) {
            Ordering::Equal => {}
            other => return other,
        }
        for pair in left.iter().zip(right.iter()) {
            match pair.0.cmp(pair.1) {
                Ordering::Equal => {}
                other => return other,
            }
        }
        Ordering::Equal
    }

    pub fn cmp_value(&self, other: &Self) -> std::cmp::Ordering {
        use std::cmp::Ordering;
        match (self.is_zero(), other.is_zero()) {
            (true, true) => return Ordering::Equal,
            (true, false) => {
                return if other.negative {
                    Ordering::Greater
                } else {
                    Ordering::Less
                }
            }
            (false, true) => {
                return if self.negative {
                    Ordering::Less
                } else {
                    Ordering::Greater
                }
            }
            (false, false) => {}
        }
        match (self.negative, other.negative) {
            (true, false) => return Ordering::Less,
            (false, true) => return Ordering::Greater,
            _ => {}
        }
        let exponent = self.exponent.min(other.exponent);
        let left = self.rescale_to(exponent);
        let right = other.rescale_to(exponent);
        let ordering = Self::magnitude_compare(&left, &right);
        if self.negative {
            ordering.reverse()
        } else {
            ordering
        }
    }

    fn add_magnitude(left: &[u8], right: &[u8]) -> Vec<u8> {
        let mut result = Vec::with_capacity(left.len().max(right.len()) + 1);
        let mut carry = 0usize;
        for index in 0..left.len().max(right.len()) {
            let left_digit = left
                .get(left.len().saturating_sub(1 + index))
                .copied()
                .unwrap_or(0) as usize;
            let right_digit = if index < right.len() {
                right[right.len() - 1 - index] as usize
            } else {
                0
            };
            let summed = left_digit + right_digit + carry;
            result.push((summed % 10) as u8);
            carry = summed / 10;
        }
        if carry > 0 {
            result.push(carry as u8);
        }
        result.reverse();
        result
    }

    fn subtract_magnitude(left: &[u8], right: &[u8]) -> Vec<u8> {
        let mut result = Vec::with_capacity(left.len());
        let mut borrow = 0i32;
        for index in 0..left.len() {
            let left_digit = left[left.len() - 1 - index] as i32 - borrow;
            let right_digit = if index < right.len() {
                right[right.len() - 1 - index] as i32
            } else {
                0
            };
            let (digit, new_borrow) = if left_digit < right_digit {
                (left_digit + 10 - right_digit, 1)
            } else {
                (left_digit - right_digit, 0)
            };
            result.push(digit as u8);
            borrow = new_borrow;
        }
        result.reverse();
        strip_leading_zeros(&mut result);
        result
    }

    fn multiply_magnitude(left: &[u8], right: &[u8]) -> Vec<u8> {
        if left.is_empty() || right.is_empty() {
            return Vec::new();
        }
        let mut result = vec![0u8; left.len() + right.len()];
        for (left_index, left_digit) in left.iter().enumerate().rev() {
            let mut carry = 0usize;
            for (right_index, right_digit) in right.iter().enumerate().rev() {
                let position = left_index + right_index + 1;
                let summed = result[position] as usize
                    + (*left_digit as usize) * (*right_digit as usize)
                    + carry;
                result[position] = (summed % 10) as u8;
                carry = summed / 10;
            }
            let mut position = left_index;
            while carry > 0 {
                let summed = result[position] as usize + carry;
                result[position] = (summed % 10) as u8;
                carry = summed / 10;
                if position == 0 {
                    break;
                }
                position -= 1;
            }
        }
        strip_leading_zeros(&mut result);
        result
    }

    fn aligned_pair(left: &Self, right: &Self) -> (Vec<u8>, Vec<u8>, bool, bool, i32) {
        let exponent = left.exponent.min(right.exponent);
        (
            left.rescale_to(exponent),
            right.rescale_to(exponent),
            left.negative,
            right.negative,
            exponent,
        )
    }

    pub fn add(&self, other: &Self) -> Self {
        let (left, right, left_negative, right_negative, exponent) =
            Self::aligned_pair(self, other);
        let digits = if left_negative == right_negative {
            Self::add_magnitude(&left, &right)
        } else {
            match Self::magnitude_compare(&left, &right) {
                std::cmp::Ordering::Equal => return Self::zero(),
                std::cmp::Ordering::Greater => {
                    let digits = Self::subtract_magnitude(&left, &right);
                    return Self {
                        negative: left_negative,
                        digits,
                        exponent,
                    }
                    .apply_context();
                }
                std::cmp::Ordering::Less => {
                    let digits = Self::subtract_magnitude(&right, &left);
                    return Self {
                        negative: right_negative,
                        digits,
                        exponent,
                    }
                    .apply_context();
                }
            }
        };
        Self {
            negative: digits.iter().any(|digit| *digit != 0) && left_negative,
            digits,
            exponent,
        }
        .apply_context()
    }

    pub fn neg(&self) -> Self {
        Self {
            negative: !self.is_zero() && !self.negative,
            digits: self.digits.clone(),
            exponent: self.exponent,
        }
    }

    pub fn multiply(&self, other: &Self) -> Self {
        if self.is_zero() || other.is_zero() {
            return Self::zero();
        }
        let digits = Self::multiply_magnitude(&self.magnitude(), &other.magnitude());
        Self {
            negative: self.negative != other.negative && !digits.iter().all(|digit| *digit == 0),
            digits,
            exponent: self.exponent + other.exponent,
        }
        .apply_context()
    }

    /// Long division normalized through scientific form: the quotient
    /// coefficient carries CONTEXT_DIGITS + 2 significant digits and is then
    /// banker-rounded to the context.
    pub fn divide(&self, other: &Self) -> Self {
        if self.is_zero() {
            return Self::zero();
        }
        let dividend = self.magnitude();
        let divisor = other.magnitude();
        let target_digits = CONTEXT_DIGITS + 2;
        let mut scaled = dividend.clone();
        scaled.extend(std::iter::repeat(0).take(target_digits));
        let scaled_len = scaled.len();
        let mut quotient: Vec<u8> = Vec::new();
        let mut remainder: Vec<u8> = Vec::new();
        let mut source_index = 0usize;
        loop {
            while remainder.first() == Some(&0) {
                remainder.remove(0);
            }
            if source_index < scaled_len {
                remainder.push(scaled[source_index]);
                source_index += 1;
            } else if remainder.is_empty() {
                break;
            } else {
                remainder.push(0);
            }
            strip_leading_zeros(&mut remainder);
            let mut digit = 0u8;
            while !remainder.is_empty()
                && Self::magnitude_compare(&remainder, &divisor) != std::cmp::Ordering::Less
            {
                remainder = Self::subtract_magnitude(&remainder, &divisor);
                digit += 1;
            }
            quotient.push(digit);
            if remainder.is_empty() {
                // Exact termination: the remaining pulled digits would only
                // append value-preserving trailing zeros, which Python's
                // ideal-exponent semantics trim.
                break;
            }
            if quotient.len() >= target_digits {
                break;
            }
        }
        let exact_termination = remainder.is_empty();
        strip_leading_zeros(&mut quotient);
        if exact_termination {
            let trailing = quotient
                .iter()
                .rev()
                .take_while(|digit| **digit == 0)
                .count();
            for _ in 0..trailing {
                quotient.pop();
                source_index -= 1;
            }
        }
        if quotient.is_empty() {
            return Self::zero();
        }
        // value = quotient_integer × 10^(self.exp - other.exp - (consumed -
        // dividend_digits)) with `consumed` the pulled prefix length.
        let exponent_base =
            self.exponent - other.exponent - (source_index as i32) + (dividend.len() as i32);
        Self::round_significant(
            quotient,
            self.negative != other.negative,
            exponent_base,
            CONTEXT_DIGITS,
        )
    }

    fn apply_context(self) -> Self {
        if self.is_zero() || self.digits.len() <= CONTEXT_DIGITS {
            return self;
        }
        let negative = self.negative;
        let exponent = self.exponent;
        let rounded =
            Self::round_significant(self.magnitude(), negative, self.exponent, CONTEXT_DIGITS);
        let _ = (negative, exponent);
        rounded
    }

    fn round_significant(digits: Vec<u8>, negative: bool, exponent: i32, context: usize) -> Self {
        let significant = digits.len();
        let keep = context.min(significant);
        let mut kept: Vec<u8> = digits[..keep].to_vec();
        if significant > context {
            let round_digit = digits[keep];
            let rest_nonzero = digits[keep + 1..].iter().any(|digit| *digit != 0);
            let round_up = round_digit > 5
                || (round_digit == 5
                    && (rest_nonzero || kept.last().is_some_and(|last| last % 2 == 1)));
            if round_up {
                let mut carry = 1usize;
                for digit in kept.iter_mut().rev() {
                    let summed = *digit as usize + carry;
                    *digit = (summed % 10) as u8;
                    carry = summed / 10;
                }
                if carry > 0 {
                    kept.insert(0, carry as u8);
                }
            }
        }
        let new_exponent = exponent + (significant - keep) as i32;
        let negative = negative && !kept.iter().all(|digit| *digit == 0);
        Self {
            negative,
            digits: kept,
            exponent: new_exponent,
        }
    }

    /// ROUND quantization with ROUND_HALF_UP on the magnitude.
    pub fn quantize_half_up(&self, digits: i32) -> Self {
        let target_exponent = -digits;
        let shift = target_exponent - self.exponent;
        if shift <= 0 {
            let mut kept = self.magnitude();
            for _ in 0..(-shift) {
                kept.push(0);
            }
            strip_leading_zeros(&mut kept);
            let negative = self.negative && !kept.iter().all(|digit| *digit == 0);
            return Self {
                negative,
                digits: kept,
                exponent: target_exponent,
            };
        }
        let magnitude = self.magnitude();
        let shift = shift as usize;
        if shift >= magnitude.len() {
            // Everything is dropped; HALF_UP keeps a 1 at the target scale only
            // when the first dropped digit is at least 5.
            let first = magnitude.first().copied().unwrap_or(0);
            if first >= 5 {
                return Self {
                    negative: self.negative,
                    digits: vec![1],
                    exponent: target_exponent,
                };
            }
            return Self::zero();
        }
        let cut = magnitude.len() - shift;
        let mut kept: Vec<u8> = magnitude[..cut].to_vec();
        if kept.is_empty() {
            kept = vec![0];
        }
        let first_dropped = magnitude[cut];
        if first_dropped >= 5 {
            let mut carry = 1usize;
            for digit in kept.iter_mut().rev() {
                let summed = *digit as usize + carry;
                *digit = (summed % 10) as u8;
                carry = summed / 10;
            }
            if carry > 0 {
                kept.insert(0, carry as u8);
            }
        }
        strip_leading_zeros(&mut kept);
        let negative = self.negative && !kept.iter().all(|digit| *digit == 0);
        Self {
            negative,
            digits: kept,
            exponent: target_exponent,
        }
    }
}

fn strip_leading_zeros(digits: &mut Vec<u8>) {
    while digits.len() > 1 && digits.first() == Some(&0) {
        digits.remove(0);
    }
    if digits.len() == 1 && digits[0] == 0 {
        digits.clear();
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum NumberValue {
    Binary64(f64),
    Decimal(Decimal28),
}

impl NumberValue {
    fn is_zero(&self) -> bool {
        match self {
            Self::Binary64(value) => *value == 0.0,
            Self::Decimal(value) => value.is_zero(),
        }
    }

    fn zero_for(numeric_mode: &str) -> Self {
        if numeric_mode == "excel_binary64" {
            Self::Binary64(0.0)
        } else {
            Self::Decimal(Decimal28::zero())
        }
    }

    fn negate(self) -> Self {
        match self {
            Self::Binary64(value) => Self::Binary64(-value),
            Self::Decimal(value) => Self::Decimal(value.neg()),
        }
    }

    fn abs_value(self) -> Self {
        match self {
            Self::Binary64(value) => Self::Binary64(value.abs()),
            Self::Decimal(value) => {
                if value.negative {
                    Self::Decimal(value.neg())
                } else {
                    Self::Decimal(value)
                }
            }
        }
    }

    fn as_decimal(&self) -> Decimal28 {
        match self {
            Self::Decimal(value) => value.clone(),
            Self::Binary64(raw) => {
                Decimal28::parse(&format!("{raw}")).unwrap_or_else(|_| Decimal28::zero())
            }
        }
    }

    fn to_integral_digits(&self) -> Result<i64, EvalError> {
        match self {
            Self::Binary64(value) => {
                if value.fract() != 0.0 || !value.is_finite() {
                    return Err(EvalError::evaluation(
                        "TYPE_ERROR：ROUND 的 digits 必须是整数",
                    ));
                }
                Ok(*value as i64)
            }
            Self::Decimal(decimal) => {
                let text = decimal.to_canonical_string();
                let integral = match text.split_once('.') {
                    Some((head, fraction)) => {
                        let head = head.trim_start_matches('-');
                        fraction.bytes().all(|byte| byte == b'0') && !head.is_empty()
                    }
                    None => true,
                };
                if !integral {
                    return Err(EvalError::evaluation(
                        "TYPE_ERROR：ROUND 的 digits 必须是整数",
                    ));
                }
                match text.parse::<i64>() {
                    Ok(digits) => Ok(digits),
                    Err(_) => Err(EvalError::evaluation(
                        "ARGUMENT_ERROR：ROUND 的 digits 范围为 -12 到 12",
                    )),
                }
            }
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum EvalScalar {
    Number(NumberValue),
    Text(String),
    Boolean(bool),
    Tuple(Vec<String>),
    Blank,
}

#[derive(Clone, Debug, PartialEq)]
pub struct EvalValue {
    pub value: EvalScalar,
    pub kind: &'static str,
    pub unit: Option<String>,
    pub metric_type: String,
    pub evidence: Vec<JsonValue>,
    pub is_literal: bool,
    pub precision: Option<i64>,
    pub numeric_mode: String,
}

impl EvalValue {
    pub fn number_from_text(text: &str, numeric_mode: &str) -> Result<Self, EvalError> {
        let number = if numeric_mode == "excel_binary64" {
            let raw: f64 = text
                .parse()
                .map_err(|_| EvalError::evaluation(format!("无效数字：{text}")))?;
            NumberValue::Binary64(excel_round15(raw))
        } else {
            NumberValue::Decimal(
                Decimal28::parse(text).map_err(|message| EvalError::evaluation(message))?,
            )
        };
        Ok(Self {
            value: EvalScalar::Number(number),
            kind: "number",
            unit: None,
            metric_type: String::new(),
            evidence: Vec::new(),
            is_literal: true,
            precision: None,
            numeric_mode: numeric_mode.to_owned(),
        })
    }

    pub fn string(value: impl Into<String>, is_literal: bool) -> Self {
        Self {
            value: EvalScalar::Text(value.into()),
            kind: "string",
            unit: None,
            metric_type: "text".to_owned(),
            evidence: Vec::new(),
            is_literal,
            precision: None,
            numeric_mode: "decimal".to_owned(),
        }
    }

    pub fn boolean(value: bool, evidence: Vec<JsonValue>, is_literal: bool) -> Self {
        Self {
            value: EvalScalar::Boolean(value),
            kind: "boolean",
            unit: None,
            metric_type: "boolean".to_owned(),
            evidence,
            is_literal,
            precision: None,
            numeric_mode: "decimal".to_owned(),
        }
    }

    fn kind_str(&self) -> &'static str {
        match &self.value {
            EvalScalar::Number(_) => "number",
            EvalScalar::Text(_) => "string",
            EvalScalar::Boolean(_) => "boolean",
            EvalScalar::Tuple(_) => "tuple",
            EvalScalar::Blank => "blank",
        }
    }
}

fn merge_evidence(values: &[&EvalValue]) -> Vec<JsonValue> {
    let mut merged: Vec<JsonValue> = Vec::new();
    let mut seen: Vec<(String, String, String, String)> = Vec::new();
    for value in values {
        for item in &value.evidence {
            let key = (
                item.get("source_id")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default()
                    .to_owned(),
                item.get("source_sha256")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default()
                    .to_owned(),
                item.get("sheet")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default()
                    .to_owned(),
                item.get("cell")
                    .and_then(JsonValue::as_str)
                    .unwrap_or_default()
                    .to_owned(),
            );
            if !seen.contains(&key) {
                seen.push(key);
                merged.push(item.clone());
            }
        }
    }
    merged
}

fn combine_precision(values: &[&EvalValue]) -> Option<i64> {
    values.iter().filter_map(|value| value.precision).max()
}

fn require_number(value: &EvalValue, function: &str) -> Result<NumberValue, EvalError> {
    match &value.value {
        EvalScalar::Blank => Ok(NumberValue::zero_for(&value.numeric_mode)),
        EvalScalar::Number(number) => Ok(number.clone()),
        _ => Err(EvalError::evaluation(format!(
            "TYPE_ERROR：{function}需要数值"
        ))),
    }
}

fn require_boolean(value: &EvalValue, function: &str) -> Result<bool, EvalError> {
    match value.value {
        EvalScalar::Boolean(flag) => Ok(flag),
        _ => Err(EvalError::evaluation(format!(
            "TYPE_ERROR：{function}只接受布尔值"
        ))),
    }
}

fn with_numeric_mode(value: EvalValue) -> EvalValue {
    if value.numeric_mode != "excel_binary64" {
        return value;
    }
    match value.value {
        EvalScalar::Number(NumberValue::Binary64(raw)) => {
            let mut updated = value;
            updated.value = EvalScalar::Number(NumberValue::Binary64(excel_round15(raw)));
            updated
        }
        EvalScalar::Blank => {
            let mut updated = value;
            updated.value = EvalScalar::Number(NumberValue::Binary64(0.0));
            updated
        }
        _ => value,
    }
}

fn unit_text(unit: &Option<String>) -> String {
    unit.clone().unwrap_or_else(|| "无单位".to_owned())
}

fn numeric_comparison_values(
    left: &EvalValue,
    right: &EvalValue,
) -> Result<(NumberValue, NumberValue), EvalError> {
    let left_number = require_number(left, "比较")?;
    let right_number = require_number(right, "比较")?;
    if left.unit == right.unit
        || (left.unit.is_none() && left.is_literal)
        || (right.unit.is_none() && right.is_literal)
    {
        // Unit rules satisfied.
    } else {
        return Err(EvalError::evaluation(format!(
            "UNIT_MISMATCH：不能比较 {} 与 {}",
            unit_text(&left.unit),
            unit_text(&right.unit)
        )));
    }
    if left.numeric_mode == "excel_binary64" || right.numeric_mode == "excel_binary64" {
        let round = |number: &NumberValue| match number {
            NumberValue::Binary64(raw) => NumberValue::Binary64(excel_round15(*raw)),
            NumberValue::Decimal(decimal) => NumberValue::Binary64(excel_round15(decimal.to_f64())),
        };
        return Ok((round(&left_number), round(&right_number)));
    }
    Ok((left_number, right_number))
}

fn compare_number_values(left: &NumberValue, right: &NumberValue) -> std::cmp::Ordering {
    match (left, right) {
        (NumberValue::Binary64(left), NumberValue::Binary64(right)) => {
            left.partial_cmp(right).unwrap_or(std::cmp::Ordering::Equal)
        }
        (NumberValue::Decimal(left), NumberValue::Decimal(right)) => left.cmp_value(right),
        (NumberValue::Binary64(left), NumberValue::Decimal(right)) => left
            .partial_cmp(&right.to_f64())
            .unwrap_or(std::cmp::Ordering::Equal),
        (NumberValue::Decimal(left), NumberValue::Binary64(right)) => left
            .to_f64()
            .partial_cmp(right)
            .unwrap_or(std::cmp::Ordering::Equal),
    }
}

fn compare_values(operator: &str, left: &EvalValue, right: &EvalValue) -> Result<bool, EvalError> {
    let left_kind = left.kind_str();
    let right_kind = right.kind_str();
    if left_kind != right_kind {
        let numeric_pair =
            matches!(left_kind, "number" | "blank") && matches!(right_kind, "number" | "blank");
        if !numeric_pair {
            return Err(EvalError::evaluation(format!(
                "TYPE_ERROR：不能比较 {left_kind} 与 {right_kind}"
            )));
        }
    }
    if matches!(left_kind, "number" | "blank") {
        let (left_number, right_number) = numeric_comparison_values(left, right)?;
        let ordering = compare_number_values(&left_number, &right_number);
        return Ok(match operator {
            "=" | "==" => ordering == std::cmp::Ordering::Equal,
            "<>" | "!=" => ordering != std::cmp::Ordering::Equal,
            "<" => ordering == std::cmp::Ordering::Less,
            "<=" => ordering != std::cmp::Ordering::Greater,
            ">" => ordering == std::cmp::Ordering::Greater,
            ">=" => ordering != std::cmp::Ordering::Less,
            _ => unreachable!("comparison operator validated by the parser"),
        });
    }
    if matches!(operator, "=" | "==" | "<>" | "!=") {
        if !matches!(left_kind, "string" | "boolean") {
            return Err(EvalError::evaluation("TYPE_ERROR：此类型只支持相等比较"));
        }
        let equal = match (&left.value, &right.value) {
            (EvalScalar::Text(left), EvalScalar::Text(right)) => left == right,
            (EvalScalar::Boolean(left), EvalScalar::Boolean(right)) => left == right,
            _ => false,
        };
        return Ok(if matches!(operator, "=" | "==") {
            equal
        } else {
            !equal
        });
    }
    Err(EvalError::evaluation(
        "TYPE_ERROR：字符串和布尔值只支持相等比较",
    ))
}

fn arithmetic_number(
    operator: &str,
    left: &NumberValue,
    right: &NumberValue,
) -> Result<NumberValue, EvalError> {
    match (left, right) {
        (NumberValue::Binary64(left), NumberValue::Binary64(right)) => {
            let value = match operator {
                "+" => left + right,
                "-" => left - right,
                "*" => left * right,
                "/" => left / right,
                _ => unreachable!("arithmetic operator validated by the parser"),
            };
            Ok(NumberValue::Binary64(value))
        }
        _ => {
            let left = left.as_decimal();
            let right = right.as_decimal();
            let value = match operator {
                "+" => left.add(&right),
                "-" => left.add(&right.neg()),
                "*" => left.multiply(&right),
                "/" => left.divide(&right),
                _ => unreachable!("arithmetic operator validated by the parser"),
            };
            Ok(NumberValue::Decimal(value))
        }
    }
}

fn combine_units(left: Option<&String>, right: Option<&String>, operator: &str) -> Option<String> {
    match operator {
        "*" => match (left, right) {
            (None, Some(right)) => Some(right.clone()),
            (Some(left), None) => Some(left.clone()),
            (Some(left), Some(right)) => Some(format!("{left}·{right}")),
            (None, None) => None,
        },
        "/" => match (left, right) {
            (Some(left), Some(right)) if left == right => None,
            (None, Some(right)) => Some(format!("1/{right}")),
            (Some(left), None) => Some(left.clone()),
            (Some(left), Some(right)) => Some(format!("{left}/{right}")),
            (None, None) => None,
        },
        _ => left.cloned(),
    }
}

fn number_eval_value(
    number: NumberValue,
    unit: Option<String>,
    metric_type: String,
    evidence: Vec<JsonValue>,
    precision: Option<i64>,
    numeric_mode: &str,
) -> EvalValue {
    let number = if numeric_mode == "excel_binary64" {
        match number {
            NumberValue::Binary64(raw) => NumberValue::Binary64(excel_round15(raw)),
            NumberValue::Decimal(decimal) => NumberValue::Binary64(excel_round15(decimal.to_f64())),
        }
    } else {
        number
    };
    EvalValue {
        value: EvalScalar::Number(number),
        kind: "number",
        unit,
        metric_type,
        evidence,
        is_literal: false,
        precision,
        numeric_mode: numeric_mode.to_owned(),
    }
}

/// Variable resolver and call-value provider injected by the rule engine.
pub trait EvaluationHooks {
    fn resolve_variable(&self, name: &str) -> Result<EvalValue, EvalError>;
    fn call_value(
        &self,
        name: &str,
        arguments: &[EvalArgument],
        keywords: &BTreeMap<String, String>,
    ) -> Result<EvalValue, EvalError>;
}

#[derive(Clone, Debug, PartialEq)]
pub enum EvalArgument {
    Text(String),
    Tuple(Vec<String>),
}

pub fn evaluate(
    node: &Node,
    hooks: &dyn EvaluationHooks,
    numeric_mode: &str,
) -> Result<EvalValue, EvalError> {
    match node {
        Node::Literal(literal) => match &literal.value {
            crate::s5_config_dsl::LiteralValue::Number(text) => {
                EvalValue::number_from_text(text, numeric_mode)
            }
            crate::s5_config_dsl::LiteralValue::String(value) => {
                Ok(EvalValue::string(value.clone(), true))
            }
            crate::s5_config_dsl::LiteralValue::Boolean(value) => {
                Ok(EvalValue::boolean(*value, Vec::new(), true))
            }
        },
        Node::Variable(variable) => Ok(with_numeric_mode(hooks.resolve_variable(&variable.name)?)),
        Node::Tuple(tuple) => Ok(EvalValue {
            value: EvalScalar::Tuple(tuple.values.clone()),
            kind: "tuple",
            unit: None,
            metric_type: String::new(),
            evidence: Vec::new(),
            is_literal: true,
            precision: None,
            numeric_mode: "decimal".to_owned(),
        }),
        Node::Unary { operator, operand } => {
            let value = evaluate(operand, hooks, numeric_mode)?;
            match operator {
                crate::s5_config_dsl::UnaryOperator::Not => Ok(EvalValue::boolean(
                    !require_boolean(&value, "NOT")?,
                    value.evidence.clone(),
                    false,
                )),
                sign => {
                    let number = require_number(&value, sign.as_str())?;
                    let number = match sign {
                        crate::s5_config_dsl::UnaryOperator::Minus => number.negate(),
                        _ => number,
                    };
                    Ok(number_eval_value(
                        number,
                        value.unit.clone(),
                        value.metric_type.clone(),
                        value.evidence.clone(),
                        value.precision,
                        &value.numeric_mode,
                    ))
                }
            }
        }
        Node::Binary {
            operator,
            left,
            right,
        } => {
            let left_value = evaluate(left, hooks, numeric_mode)?;
            match operator {
                crate::s5_config_dsl::BinaryOperator::And => {
                    if !require_boolean(&left_value, "AND")? {
                        return Ok(EvalValue::boolean(
                            false,
                            left_value.evidence.clone(),
                            false,
                        ));
                    }
                    let right_value = evaluate(right, hooks, numeric_mode)?;
                    Ok(EvalValue::boolean(
                        require_boolean(&right_value, "AND")?,
                        merge_evidence(&[&left_value, &right_value]),
                        false,
                    ))
                }
                crate::s5_config_dsl::BinaryOperator::Or => {
                    if require_boolean(&left_value, "OR")? {
                        return Ok(EvalValue::boolean(true, left_value.evidence.clone(), false));
                    }
                    let right_value = evaluate(right, hooks, numeric_mode)?;
                    Ok(EvalValue::boolean(
                        require_boolean(&right_value, "OR")?,
                        merge_evidence(&[&left_value, &right_value]),
                        false,
                    ))
                }
                comparison if comparison.is_comparison() => {
                    let right_value = evaluate(right, hooks, numeric_mode)?;
                    Ok(EvalValue::boolean(
                        compare_values(comparison.as_str(), &left_value, &right_value)?,
                        merge_evidence(&[&left_value, &right_value]),
                        false,
                    ))
                }
                arithmetic => {
                    let right_value = evaluate(right, hooks, numeric_mode)?;
                    let operator_text = arithmetic.as_str();
                    if matches!(
                        arithmetic,
                        crate::s5_config_dsl::BinaryOperator::Add
                            | crate::s5_config_dsl::BinaryOperator::Subtract
                    ) {
                        let left_is_unitless_literal =
                            left_value.unit.is_none() && left_value.is_literal;
                        let right_is_unitless_literal =
                            right_value.unit.is_none() && right_value.is_literal;
                        if left_value.unit != right_value.unit
                            && !left_is_unitless_literal
                            && !right_is_unitless_literal
                        {
                            return Err(EvalError::evaluation(format!(
                                "UNIT_MISMATCH：不能进行 {operator_text} 运算：{} / {}",
                                unit_text(&left_value.unit),
                                unit_text(&right_value.unit)
                            )));
                        }
                    }
                    let left_number = require_number(&left_value, operator_text)?;
                    let right_number = require_number(&right_value, operator_text)?;
                    let result_mode = if left_value.numeric_mode == "excel_binary64"
                        || right_value.numeric_mode == "excel_binary64"
                    {
                        "excel_binary64"
                    } else {
                        "decimal"
                    };
                    if matches!(
                        arithmetic,
                        crate::s5_config_dsl::BinaryOperator::Multiply
                            | crate::s5_config_dsl::BinaryOperator::Divide
                    ) {
                        let unit = combine_units(
                            left_value.unit.as_ref(),
                            right_value.unit.as_ref(),
                            operator_text,
                        );
                        let number = match arithmetic {
                            crate::s5_config_dsl::BinaryOperator::Divide => {
                                if right_number.is_zero() {
                                    return Err(EvalError::evaluation(
                                        "DIVISION_BY_ZERO：除数为零",
                                    ));
                                }
                                arithmetic_number("/", &left_number, &right_number)?
                            }
                            _ => arithmetic_number("*", &left_number, &right_number)?,
                        };
                        return Ok(number_eval_value(
                            number,
                            unit,
                            "derived".to_owned(),
                            merge_evidence(&[&left_value, &right_value]),
                            combine_precision(&[&left_value, &right_value]),
                            result_mode,
                        ));
                    }
                    let number = arithmetic_number(operator_text, &left_number, &right_number)?;
                    let metric_type = if left_value.metric_type.is_empty() {
                        right_value.metric_type.clone()
                    } else {
                        left_value.metric_type.clone()
                    };
                    Ok(number_eval_value(
                        number,
                        left_value.unit.clone().or(right_value.unit.clone()),
                        metric_type,
                        merge_evidence(&[&left_value, &right_value]),
                        combine_precision(&[&left_value, &right_value]),
                        result_mode,
                    ))
                }
            }
        }
        Node::Call(call) => {
            let name = call.name.as_str();
            match name {
                "IF" => {
                    let condition = evaluate(&call.arguments[0], hooks, numeric_mode)?;
                    let branch = if require_boolean(&condition, "IF")? {
                        &call.arguments[1]
                    } else {
                        &call.arguments[2]
                    };
                    let result = evaluate(branch, hooks, numeric_mode)?;
                    Ok(EvalValue {
                        value: result.value.clone(),
                        kind: result.kind_str(),
                        unit: result.unit.clone(),
                        metric_type: result.metric_type.clone(),
                        evidence: merge_evidence(&[&condition, &result]),
                        is_literal: result.is_literal,
                        precision: result.precision,
                        numeric_mode: result.numeric_mode.clone(),
                    })
                }
                "AND" | "OR" => {
                    let is_and = name == "AND";
                    let mut evaluated: Vec<EvalValue> = Vec::new();
                    for argument in &call.arguments {
                        let value = evaluate(argument, hooks, numeric_mode)?;
                        let flag = require_boolean(&value, name)?;
                        let mut seen: Vec<&EvalValue> = evaluated.iter().collect();
                        seen.push(&value);
                        if is_and && !flag {
                            return Ok(EvalValue::boolean(false, merge_evidence(&seen), false));
                        }
                        if !is_and && flag {
                            return Ok(EvalValue::boolean(true, merge_evidence(&seen), false));
                        }
                        evaluated.push(value);
                    }
                    let seen: Vec<&EvalValue> = evaluated.iter().collect();
                    Ok(EvalValue::boolean(is_and, merge_evidence(&seen), false))
                }
                "NOT" => {
                    let value = evaluate(&call.arguments[0], hooks, numeric_mode)?;
                    Ok(EvalValue::boolean(
                        !require_boolean(&value, "NOT")?,
                        value.evidence.clone(),
                        false,
                    ))
                }
                "V_CUR" | "V_PRE" | "V_EXT" => {
                    let mut arguments: Vec<EvalArgument> = Vec::new();
                    for argument in &call.arguments {
                        match argument {
                            Node::Literal(literal) => match &literal.value {
                                crate::s5_config_dsl::LiteralValue::String(value) => {
                                    arguments.push(EvalArgument::Text(value.clone()));
                                }
                                crate::s5_config_dsl::LiteralValue::Number(value) => {
                                    arguments.push(EvalArgument::Text(value.clone()));
                                }
                                crate::s5_config_dsl::LiteralValue::Boolean(value) => {
                                    arguments.push(EvalArgument::Text(value.to_string()));
                                }
                            },
                            Node::Tuple(tuple) => {
                                arguments.push(EvalArgument::Tuple(tuple.values.clone()));
                            }
                            _ => {
                                return Err(EvalError::configuration(format!(
                                    "函数 {name} 的定位参数必须是常量"
                                )));
                            }
                        }
                    }
                    let keywords: BTreeMap<String, String> =
                        call.keywords.iter().cloned().collect();
                    Ok(with_numeric_mode(
                        hooks.call_value(name, &arguments, &keywords)?,
                    ))
                }
                _ => {
                    let values: Vec<EvalValue> = call
                        .arguments
                        .iter()
                        .map(|argument| evaluate(argument, hooks, numeric_mode))
                        .collect::<Result<Vec<_>, _>>()?;
                    match name {
                        "ISBLANK" => Ok(EvalValue::boolean(
                            values[0].kind_str() == "blank",
                            values[0].evidence.clone(),
                            false,
                        )),
                        "ISNUMBER" => Ok(EvalValue::boolean(
                            values[0].kind_str() == "number",
                            values[0].evidence.clone(),
                            false,
                        )),
                        "SUM" => {
                            for value in &values {
                                if !matches!(value.kind_str(), "number" | "blank" | "string") {
                                    return Err(EvalError::evaluation(
                                        "TYPE_ERROR：SUM 只接受数值、文本或空值",
                                    ));
                                }
                            }
                            let numeric_positions: Vec<usize> = values
                                .iter()
                                .enumerate()
                                .filter(|(_, value)| value.kind_str() == "number")
                                .map(|(index, _)| index)
                                .collect();
                            let measured_position = values
                                .iter()
                                .position(|value| value.unit.is_some() && !value.is_literal);
                            let unit_source_position =
                                measured_position.or(numeric_positions.first().copied());
                            let unit =
                                unit_source_position.and_then(|index| values[index].unit.clone());
                            for (index, value) in values.iter().enumerate() {
                                if Some(index) == unit_source_position {
                                    continue;
                                }
                                if !matches!(value.kind_str(), "number" | "blank") {
                                    continue;
                                }
                                if value.unit != unit && !(value.unit.is_none() && value.is_literal)
                                {
                                    return Err(EvalError::evaluation(
                                        "UNIT_MISMATCH：SUM 的参数单位不一致",
                                    ));
                                }
                            }
                            let result_mode = if values
                                .iter()
                                .any(|value| value.numeric_mode == "excel_binary64")
                            {
                                "excel_binary64"
                            } else {
                                "decimal"
                            };
                            let mut total = NumberValue::zero_for(result_mode);
                            for index in &numeric_positions {
                                let number = require_number(&values[*index], "SUM")?;
                                total = arithmetic_number("+", &total, &number)?;
                            }
                            let metric_type = numeric_positions
                                .first()
                                .map(|index| values[*index].metric_type.clone())
                                .or_else(|| {
                                    unit_source_position
                                        .map(|index| values[index].metric_type.clone())
                                })
                                .unwrap_or_default();
                            let references: Vec<&EvalValue> = values.iter().collect();
                            Ok(number_eval_value(
                                total,
                                unit,
                                metric_type,
                                merge_evidence(&references),
                                combine_precision(&references),
                                result_mode,
                            ))
                        }
                        "ABS" => {
                            let value = &values[0];
                            let number = require_number(value, "ABS")?;
                            Ok(number_eval_value(
                                number.abs_value(),
                                value.unit.clone(),
                                value.metric_type.clone(),
                                value.evidence.clone(),
                                value.precision,
                                &value.numeric_mode,
                            ))
                        }
                        "ROUND" => {
                            let value = &values[0];
                            let digits =
                                require_number(&values[1], "ROUND")?.to_integral_digits()?;
                            if digits.abs() > 12 {
                                return Err(EvalError::evaluation(
                                    "ARGUMENT_ERROR：ROUND 的 digits 范围为 -12 到 12",
                                ));
                            }
                            let number = require_number(value, "ROUND")?;
                            let rounded = match &number {
                                NumberValue::Decimal(decimal) => {
                                    NumberValue::Decimal(decimal.quantize_half_up(digits as i32))
                                }
                                NumberValue::Binary64(raw) => {
                                    let decimal = Decimal28::parse(&format!("{raw}"))
                                        .map_err(EvalError::evaluation)?;
                                    NumberValue::Binary64(excel_round15(
                                        decimal.quantize_half_up(digits as i32).to_f64(),
                                    ))
                                }
                            };
                            Ok(number_eval_value(
                                rounded,
                                value.unit.clone(),
                                value.metric_type.clone(),
                                value.evidence.clone(),
                                Some(digits.max(0)),
                                &value.numeric_mode,
                            ))
                        }
                        "MIN" | "MAX" => {
                            let is_min = name == "MIN";
                            let numbers: Vec<NumberValue> = values
                                .iter()
                                .map(|value| require_number(value, name))
                                .collect::<Result<Vec<_>, _>>()?;
                            let unit = values[0].unit.clone();
                            for value in values.iter().skip(1) {
                                if value.unit != unit
                                    && !(value.unit.is_none() && value.is_literal)
                                    && !(unit.is_none() && values[0].is_literal)
                                {
                                    return Err(EvalError::evaluation(format!(
                                        "UNIT_MISMATCH：{name} 的参数单位不一致"
                                    )));
                                }
                            }
                            let mut picked = 0usize;
                            for index in 1..numbers.len() {
                                let ordering =
                                    compare_number_values(&numbers[index], &numbers[picked]);
                                let replaces = if is_min {
                                    ordering == std::cmp::Ordering::Less
                                } else {
                                    ordering == std::cmp::Ordering::Greater
                                };
                                if replaces {
                                    picked = index;
                                }
                            }
                            let chosen = &values[picked];
                            let references: Vec<&EvalValue> = values.iter().collect();
                            Ok(number_eval_value(
                                numbers[picked].clone(),
                                chosen.unit.clone().or(unit),
                                chosen.metric_type.clone(),
                                merge_evidence(&references),
                                combine_precision(&references),
                                &chosen.numeric_mode,
                            ))
                        }
                        _ => Err(EvalError::configuration(format!("不支持的函数：{name}"))),
                    }
                }
            }
        }
    }
}

/// Convenience wrapper mirroring `evaluate_expression` with map-based hooks.
pub fn evaluate_with_maps(
    node: &Node,
    variables: &BTreeMap<String, EvalValue>,
    call_results: &BTreeMap<String, EvalValue>,
    numeric_mode: &str,
) -> Result<EvalValue, EvalError> {
    struct Maps<'a> {
        variables: &'a BTreeMap<String, EvalValue>,
        call_results: &'a BTreeMap<String, EvalValue>,
    }
    impl<'a> EvaluationHooks for Maps<'a> {
        fn resolve_variable(&self, name: &str) -> Result<EvalValue, EvalError> {
            self.variables.get(name).cloned().ok_or_else(|| {
                EvalError::evaluation(format!("UNDEFINED_VARIABLE：未定义变量 {name}"))
            })
        }
        fn call_value(
            &self,
            name: &str,
            arguments: &[EvalArgument],
            keywords: &BTreeMap<String, String>,
        ) -> Result<EvalValue, EvalError> {
            let mut key = name.to_owned();
            for argument in arguments {
                key.push('|');
                match argument {
                    EvalArgument::Text(value) => key.push_str(value),
                    EvalArgument::Tuple(values) => {
                        key.push_str(&values.join("@"));
                    }
                }
            }
            for (key_name, value) in keywords {
                key.push('|');
                key.push_str(key_name);
                key.push('=');
                key.push_str(value);
            }
            self.call_results
                .get(&key)
                .cloned()
                .ok_or_else(|| EvalError::evaluation(format!("未配置取值函数：{name}")))
        }
    }
    let hooks = Maps {
        variables,
        call_results,
    };
    evaluate(node, &hooks, numeric_mode)
}

/// Serialize an evaluated value the way the parity harness compares it.
pub fn eval_value_to_json(value: &EvalValue) -> JsonValue {
    let scalar = match &value.value {
        EvalScalar::Number(NumberValue::Binary64(raw)) => {
            // The Python baseline stores ROUND results as Decimal and floats
            // otherwise; the parity harness compares the value through a
            // float round-trip on both sides, so emit the shortest form.
            let text = if raw.fract() == 0.0 && raw.abs() < 1e16 {
                format!("{raw:.1}")
            } else {
                format!("{raw}")
            };
            json!({"kind": "number", "value": text})
        }
        EvalScalar::Number(NumberValue::Decimal(decimal)) => json!({
            "kind": "number",
            "value": decimal.to_canonical_string(),
        }),
        EvalScalar::Text(text) => json!({"kind": "string", "value": text}),
        EvalScalar::Boolean(flag) => json!({"kind": "boolean", "value": flag}),
        EvalScalar::Tuple(values) => json!({"kind": "tuple", "value": values}),
        EvalScalar::Blank => json!({"kind": "blank", "value": null}),
    };
    json!({
        "value": scalar,
        "kind": value.kind,
        "unit": value.unit,
        "metric_type": value.metric_type,
        "is_literal": value.is_literal,
        "precision": value.precision,
        "numeric_mode": value.numeric_mode,
    })
}

#[cfg(test)]
mod eval_tests {
    use super::*;
    use crate::s5_config_dsl::parse_expression;

    #[test]
    fn decimal_division_matches_python_context() {
        let one = Decimal28::from_i64(1);
        let three = Decimal28::from_i64(3);
        assert_eq!(
            one.divide(&three).to_canonical_string(),
            "0.3333333333333333333333333333"
        );
        let two = Decimal28::from_i64(2);
        assert_eq!(one.divide(&two).to_canonical_string(), "0.5");
        let a = Decimal28::parse("0.1").expect("parse");
        let b = Decimal28::parse("0.2").expect("parse");
        assert_eq!(a.add(&b).to_canonical_string(), "0.3");
        assert_eq!(a.multiply(&b).to_canonical_string(), "0.02");
    }

    #[test]
    fn decimal_quantize_is_half_up() {
        let value = Decimal28::parse("2.675").expect("parse");
        assert_eq!(value.quantize_half_up(2).to_canonical_string(), "2.68");
        let half = Decimal28::parse("2.5").expect("parse");
        assert_eq!(half.quantize_half_up(0).to_canonical_string(), "3");
        let negative_half = Decimal28::parse("-2.5").expect("parse");
        assert_eq!(
            negative_half.quantize_half_up(0).to_canonical_string(),
            "-3"
        );
        let recurring = Decimal28::from_i64(1).divide(&Decimal28::from_i64(3));
        assert_eq!(
            recurring.quantize_half_up(5).to_canonical_string(),
            "0.33333"
        );
    }

    #[test]
    fn evaluates_boolean_and_arithmetic_semantics() {
        let node = parse_expression("(NOT(ISBLANK({本期值})) AND ({本期值} <> 0))")
            .expect("parse real-shaped expression");
        let mut variables = BTreeMap::new();
        variables.insert(
            "本期值".to_owned(),
            EvalValue::number_from_text("3.5", "excel_binary64").expect("number"),
        );
        let outcome = evaluate_with_maps(
            &node,
            &variables,
            &std::collections::BTreeMap::new(),
            "excel_binary64",
        )
        .expect("evaluate");
        assert_eq!(outcome.value, EvalScalar::Boolean(true));

        let node = parse_expression("ROUND(2.675, 2)").expect("parse");
        let outcome = evaluate_with_maps(
            &node,
            &std::collections::BTreeMap::new(),
            &std::collections::BTreeMap::new(),
            "decimal",
        )
        .expect("evaluate");
        match outcome.value {
            EvalScalar::Number(NumberValue::Decimal(decimal)) => {
                assert_eq!(decimal.to_canonical_string(), "2.68");
            }
            other => panic!("unexpected ROUND result {other:?}"),
        }
    }
}
