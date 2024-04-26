use anyhow::{Result, Error};
use calamine::{Range, Data};
use serde_json::{Value, Number};
use crate::utils::conversions;

pub fn extract_cell_value(sheet: &Range<Data>, row: u32, col: u32) -> Result<Option<Value>, Error> {
    if let Some(cell) = sheet.get_value((row as u32, col as u32)) {
        match cell {
            Data::Empty => Ok(Some(Value::Null)),
            Data::Int(int_val) => Ok(Some(Value::Number(Number::from(*int_val)))),
            Data::Float(float_val) => Number::from_f64(*float_val)
                                           .map(Value::Number)
                                           .map(Some)
                                           .ok_or_else(|| Error::msg("Invalid float value")),
            Data::Bool(bool_val) => Ok(Some(Value::Bool(*bool_val))),
            Data::String(str_val) => {
                let trimmed_str = str_val.trim().to_string();
                Ok(Some(Value::String(trimmed_str)))
            },
            Data::Error(_) => Err(Error::msg("Error in cell")),
            Data::DateTime(dt) => Ok(Some(Value::String(conversions::excel_datetime(dt.as_f64())?))),
            Data::DurationIso(duration_iso) => Ok(Some(Value::String(duration_iso.to_string()))),
            _ => Err(Error::msg("Unsupported data type"))
        }
    } else {
        Ok(None)
    }
}