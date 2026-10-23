//! Exact decimal parsing and unit scaling for S2 amount values.
//!
//! This module deliberately handles only parsing and powers-of-ten scaling.
//! It does not implement general decimal arithmetic or connect to the S2
//! importer/external comparison workflow.

use crate::discovery::decimal_digit;
use excel_reader::CellValue;
use serde::{Deserialize, Serialize};
use std::fmt;

pub const AMOUNT_DATA_TYPES: &[&str] = &["余额", "累发"];

const UNIT_POWERS_OF_TEN: &[(&str, i64)] = &[("元", 0), ("万元", 4), ("亿元", 8)];
const MAX_EXACT_XLSX_INTEGER: f64 = 9_007_199_254_740_992.0;

/// Source-format normalization used by the existing S2 importer before its
/// Decimal conversion. `.xls` values follow xlrd-compatible numeric semantics;
/// OOXML values follow the source project's openpyxl-compatible normalization.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AmountSourceFormat {
    Xls,
    XlsxLike,
}

/// A finite decimal represented by the same sign/digits/exponent tuple as
/// Python `Decimal.as_tuple()`. The digit vector is arbitrary precision and
/// trailing zeroes are retained. No binary floating-point conversion occurs.
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct DecimalTuple {
    pub sign: u8,
    pub digits: Vec<u8>,
    pub exponent: i64,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum UnitConversionError {
    InvalidAmount(String),
    NonFiniteAmount(String),
    UnsupportedUnit(String),
    ExponentOutOfRange(String),
}

impl fmt::Display for UnitConversionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidAmount(value) => {
                write!(formatter, "金额不是有效数值：{}", python_string_repr(value))
            }
            Self::NonFiniteAmount(value) => {
                write!(formatter, "金额不是有限数值：{}", python_string_repr(value))
            }
            Self::UnsupportedUnit(unit) => write!(
                formatter,
                "不支持的金额单位：{}（支持：元、万元、亿元）",
                if unit.is_empty() { "空白" } else { unit }
            ),
            Self::ExponentOutOfRange(value) => write!(
                formatter,
                "金额指数超出 Rust 表示范围：{}",
                python_string_repr(value)
            ),
        }
    }
}

impl std::error::Error for UnitConversionError {}

pub fn is_amount_data_type(value: &str) -> bool {
    AMOUNT_DATA_TYPES.contains(&value)
}

/// Parse decimal text using the source helper's trim/comma-removal behavior.
/// Underscores accepted by Python's `Decimal` text parser are also ignored.
pub fn as_decimal(value: &str) -> Result<DecimalTuple, UnitConversionError> {
    parse_decimal_text(value).map_err(|error| match error {
        DecimalTextError::Invalid => UnitConversionError::InvalidAmount(value.to_owned()),
        DecimalTextError::NonFinite => UnitConversionError::NonFiniteAmount(value.to_owned()),
        DecimalTextError::ExponentOutOfRange => {
            UnitConversionError::ExponentOutOfRange(value.to_owned())
        }
    })
}

/// Scale a decimal between 元/万元/亿元 by shifting its base-ten exponent.
/// The coefficient, sign and trailing zeroes remain exact; the function never
/// applies a precision context or rounds.
pub fn convert_amount(
    value: &str,
    source_unit: &str,
    target_unit: &str,
) -> Result<DecimalTuple, UnitConversionError> {
    let source = source_unit.trim();
    let target = target_unit.trim();
    let source_power = unit_power(source)
        .ok_or_else(|| UnitConversionError::UnsupportedUnit(source.to_owned()))?;
    let target_power = unit_power(target)
        .ok_or_else(|| UnitConversionError::UnsupportedUnit(target.to_owned()))?;

    let mut number = as_decimal(value)?;
    let shift = source_power - target_power;
    number.exponent = number
        .exponent
        .checked_add(shift)
        .ok_or_else(|| UnitConversionError::ExponentOutOfRange(value.to_owned()))?;
    Ok(number)
}

/// Compatibility alias matching the source project's `to_yuan` helper.
pub fn to_yuan(value: &str, unit: &str) -> Result<DecimalTuple, UnitConversionError> {
    convert_amount(value, unit, "元")
}

/// Interpret a typed reader cell like the source S2 importer's `_decimal_number`.
/// Empty, boolean (for OOXML), date, error, invalid-text and non-finite values are
/// skipped. For `.xls`, booleans become 0/1 and integer cells pass through f64,
/// matching the legacy xlrd-facing normalization. For XLSX-like input, integral
/// floats within 2^53 become integers before Decimal parsing.
pub fn decimal_number_from_cell(
    value: Option<&CellValue>,
    source_format: AmountSourceFormat,
) -> Result<Option<DecimalTuple>, UnitConversionError> {
    Ok(decimal_from_cell(value, source_format)?.map(|(number, _)| number))
}

/// Convert a typed amount cell while preserving the source importer's skip
/// behavior for nonnumeric values. Unit validation occurs only after a numeric
/// amount exists, as it does in the S2 importer.
pub fn convert_amount_from_cell(
    value: Option<&CellValue>,
    source_format: AmountSourceFormat,
    source_unit: &str,
    target_unit: &str,
) -> Result<Option<DecimalTuple>, UnitConversionError> {
    let Some((number, original)) = decimal_from_cell(value, source_format)? else {
        return Ok(None);
    };
    let source = source_unit.trim();
    let target = target_unit.trim();
    let source_power = unit_power(source)
        .ok_or_else(|| UnitConversionError::UnsupportedUnit(source.to_owned()))?;
    let target_power = unit_power(target)
        .ok_or_else(|| UnitConversionError::UnsupportedUnit(target.to_owned()))?;
    shift_decimal(number, source_power, target_power, &original).map(Some)
}

fn decimal_from_cell(
    value: Option<&CellValue>,
    source_format: AmountSourceFormat,
) -> Result<Option<(DecimalTuple, String)>, UnitConversionError> {
    let Some(value) = value else {
        return Ok(None);
    };
    let text = match value {
        CellValue::Empty => return Ok(None),
        CellValue::Text(value) if value.is_empty() => return Ok(None),
        CellValue::Text(value) => value.clone(),
        CellValue::Integer(value) => match source_format {
            AmountSourceFormat::Xls => format!("{:?}", *value as f64),
            AmountSourceFormat::XlsxLike => value.to_string(),
        },
        CellValue::Number(value) => match source_format {
            AmountSourceFormat::Xls => format!("{value:?}"),
            AmountSourceFormat::XlsxLike
                if value.is_finite()
                    && value.fract() == 0.0
                    && value.abs() <= MAX_EXACT_XLSX_INTEGER =>
            {
                (*value as i64).to_string()
            }
            AmountSourceFormat::XlsxLike => format!("{value:?}"),
        },
        CellValue::Boolean(value) if source_format == AmountSourceFormat::Xls => {
            u8::from(*value).to_string()
        }
        CellValue::Boolean(_)
        | CellValue::DateTimeSerial(_)
        | CellValue::DateTimeIso(_)
        | CellValue::DurationIso(_)
        | CellValue::Error(_) => return Ok(None),
    };

    match as_decimal(&text) {
        Ok(number) => Ok(Some((number, text))),
        Err(UnitConversionError::InvalidAmount(_) | UnitConversionError::NonFiniteAmount(_)) => {
            Ok(None)
        }
        Err(error @ UnitConversionError::ExponentOutOfRange(_)) => Err(error),
        Err(error @ UnitConversionError::UnsupportedUnit(_)) => Err(error),
    }
}

fn unit_power(unit: &str) -> Option<i64> {
    UNIT_POWERS_OF_TEN
        .iter()
        .find_map(|(name, power)| (*name == unit).then_some(*power))
}

fn shift_decimal(
    mut number: DecimalTuple,
    source_power: i64,
    target_power: i64,
    original: &str,
) -> Result<DecimalTuple, UnitConversionError> {
    let shift = source_power - target_power;
    number.exponent = number
        .exponent
        .checked_add(shift)
        .ok_or_else(|| UnitConversionError::ExponentOutOfRange(original.to_owned()))?;
    Ok(number)
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum DecimalTextError {
    Invalid,
    NonFinite,
    ExponentOutOfRange,
}

fn parse_decimal_text(original: &str) -> Result<DecimalTuple, DecimalTextError> {
    let cleaned = original.trim().replace(',', "").replace('_', "");
    if is_non_finite(&cleaned) {
        return Err(DecimalTextError::NonFinite);
    }

    let (mantissa, explicit_exponent_text) = split_exponent(&cleaned)?;

    let (sign, unsigned_mantissa) = if let Some(unsigned) = mantissa.strip_prefix('-') {
        (1, unsigned)
    } else if let Some(unsigned) = mantissa.strip_prefix('+') {
        (0, unsigned)
    } else {
        (0, mantissa)
    };

    let mut digits = Vec::with_capacity(unsigned_mantissa.len());
    let mut fractional_digits = 0_i64;
    let mut after_decimal_point = false;
    let mut saw_decimal_point = false;

    for character in unsigned_mantissa.chars() {
        if let Some(digit) = decimal_digit(character) {
            digits.push(digit as u8);
            if after_decimal_point {
                fractional_digits = fractional_digits
                    .checked_add(1)
                    .ok_or(DecimalTextError::ExponentOutOfRange)?;
            }
        } else {
            match character {
                '.' if !saw_decimal_point => {
                    saw_decimal_point = true;
                    after_decimal_point = true;
                }
                _ => return Err(DecimalTextError::Invalid),
            }
        }
    }

    if digits.is_empty() {
        return Err(DecimalTextError::Invalid);
    }

    let explicit_exponent = explicit_exponent_text
        .map(parse_exponent)
        .transpose()?
        .unwrap_or(0);
    let exponent = explicit_exponent
        .checked_sub(fractional_digits)
        .ok_or(DecimalTextError::ExponentOutOfRange)?;
    let first_nonzero = digits
        .iter()
        .position(|digit| *digit != 0)
        .unwrap_or(digits.len().saturating_sub(1));
    if first_nonzero > 0 {
        digits.drain(..first_nonzero);
    }

    Ok(DecimalTuple {
        sign,
        digits,
        exponent,
    })
}

fn parse_exponent(value: &str) -> Result<i64, DecimalTextError> {
    let mut ascii = String::with_capacity(value.len());
    for (index, character) in value.chars().enumerate() {
        if index == 0 && matches!(character, '+' | '-') {
            ascii.push(character);
        } else if let Some(digit) = decimal_digit(character) {
            ascii.push(char::from(b'0' + digit as u8));
        } else {
            return Err(DecimalTextError::Invalid);
        }
    }
    if ascii == "+" || ascii == "-" || ascii.is_empty() {
        return Err(DecimalTextError::Invalid);
    }
    ascii
        .parse::<i64>()
        .map_err(|_| DecimalTextError::ExponentOutOfRange)
}

fn split_exponent(value: &str) -> Result<(&str, Option<&str>), DecimalTextError> {
    let mut marker = None;
    for (index, byte) in value.bytes().enumerate() {
        if byte == b'e' || byte == b'E' {
            if marker.replace(index).is_some() {
                return Err(DecimalTextError::Invalid);
            }
        }
    }

    match marker {
        Some(index) => {
            let mantissa = &value[..index];
            let exponent = &value[index + 1..];
            if exponent.is_empty() {
                return Err(DecimalTextError::Invalid);
            }
            Ok((mantissa, Some(exponent)))
        }
        None => Ok((value, None)),
    }
}

fn is_non_finite(value: &str) -> bool {
    let unsigned = value
        .strip_prefix('+')
        .or_else(|| value.strip_prefix('-'))
        .unwrap_or(value)
        .to_ascii_lowercase();
    if matches!(unsigned.as_str(), "inf" | "infinity") {
        return true;
    }

    for prefix in ["nan", "snan"] {
        if let Some(payload) = unsigned.strip_prefix(prefix) {
            return payload.bytes().all(|byte| byte.is_ascii_digit());
        }
    }
    false
}

fn python_string_repr(value: &str) -> String {
    let quote = if !value.contains('\'') || value.contains('"') {
        '\''
    } else {
        '"'
    };
    let mut result = String::with_capacity(value.len() + 2);
    result.push(quote);
    for character in value.chars() {
        match character {
            '\\' => result.push_str("\\\\"),
            '\n' => result.push_str("\\n"),
            '\r' => result.push_str("\\r"),
            '\t' => result.push_str("\\t"),
            current if current == quote => {
                result.push('\\');
                result.push(current);
            }
            current => result.push(current),
        }
    }
    result.push(quote);
    result
}

#[cfg(test)]
mod tests {
    use super::{
        as_decimal, convert_amount, convert_amount_from_cell, decimal_number_from_cell,
        AmountSourceFormat, DecimalTuple,
    };
    use excel_reader::CellValue;

    #[test]
    fn unit_scaling_preserves_arbitrary_precision_and_decimal_scale() {
        assert_eq!(
            convert_amount("123456789012345678901234567890.12345678", "元", "万元").unwrap(),
            DecimalTuple {
                sign: 0,
                digits: vec![
                    1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 0, 1, 2, 3, 4, 5, 6,
                    7, 8, 9, 0, 1, 2, 3, 4, 5, 6, 7, 8
                ],
                exponent: -12,
            }
        );
        assert_eq!(
            convert_amount("-0.00100", "万元", "元").unwrap(),
            DecimalTuple {
                sign: 1,
                digits: vec![1, 0, 0],
                exponent: -1,
            }
        );
    }

    #[test]
    fn source_unit_and_target_unit_are_trimmed() {
        assert_eq!(convert_amount("1", " 万元 ", " 元 ").unwrap().exponent, 4);
    }

    #[test]
    fn amount_parser_keeps_negative_zero_and_rejects_non_finite_values() {
        assert_eq!(
            as_decimal("-0.000").unwrap(),
            DecimalTuple {
                sign: 1,
                digits: vec![0],
                exponent: -3,
            }
        );
        assert_eq!(
            as_decimal("NaN").unwrap_err().to_string(),
            "金额不是有限数值：'NaN'"
        );
    }

    #[test]
    fn typed_cell_values_follow_source_format_normalization() {
        let integral_float = CellValue::Number(42.0);
        assert_eq!(
            decimal_number_from_cell(Some(&integral_float), AmountSourceFormat::XlsxLike).unwrap(),
            Some(DecimalTuple {
                sign: 0,
                digits: vec![4, 2],
                exponent: 0,
            })
        );
        assert_eq!(
            decimal_number_from_cell(Some(&integral_float), AmountSourceFormat::Xls).unwrap(),
            Some(DecimalTuple {
                sign: 0,
                digits: vec![4, 2, 0],
                exponent: -1,
            })
        );

        let integer = CellValue::Integer(42);
        assert_eq!(
            decimal_number_from_cell(Some(&integer), AmountSourceFormat::Xls).unwrap(),
            Some(DecimalTuple {
                sign: 0,
                digits: vec![4, 2, 0],
                exponent: -1,
            })
        );

        let boolean = CellValue::Boolean(true);
        assert_eq!(
            decimal_number_from_cell(Some(&boolean), AmountSourceFormat::XlsxLike).unwrap(),
            None
        );
        assert_eq!(
            decimal_number_from_cell(Some(&boolean), AmountSourceFormat::Xls).unwrap(),
            Some(DecimalTuple {
                sign: 0,
                digits: vec![1],
                exponent: 0,
            })
        );
    }

    #[test]
    fn typed_cell_conversion_skips_non_numeric_values_before_unit_validation() {
        for value in [
            CellValue::Empty,
            CellValue::Text(String::new()),
            CellValue::Text("not a number".to_owned()),
            CellValue::Boolean(false),
            CellValue::DateTimeSerial(45_292.0),
            CellValue::Error("Value".to_owned()),
            CellValue::Number(f64::NAN),
        ] {
            assert_eq!(
                convert_amount_from_cell(
                    Some(&value),
                    AmountSourceFormat::XlsxLike,
                    "unsupported",
                    "unknown"
                )
                .unwrap(),
                None
            );
        }

        assert_eq!(
            convert_amount_from_cell(None, AmountSourceFormat::XlsxLike, "unsupported", "unknown")
                .unwrap(),
            None
        );

        let amount = CellValue::Number(1_234_567.89);
        assert_eq!(
            convert_amount_from_cell(Some(&amount), AmountSourceFormat::XlsxLike, "元", "万元")
                .unwrap(),
            Some(DecimalTuple {
                sign: 0,
                digits: vec![1, 2, 3, 4, 5, 6, 7, 8, 9],
                exponent: -6,
            })
        );
    }
}
