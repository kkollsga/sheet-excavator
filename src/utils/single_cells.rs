use anyhow::{Result, Error};
use calamine::{Range, Data};
use serde_json::{Map, Value, Number};
use crate::utils::conversions;


fn extract_cell_value(sheet: &Range<Data>, row: u32, col: u32) -> Result<Option<Value>, Error> {
    if let Some(cell) = sheet.get_value((row as u32, col as u32)) {
        match cell {
            Data::Empty => Ok(Some(Value::Null)),
            Data::Int(int_val) => Ok(Some(Value::Number(Number::from(*int_val)))),
            Data::Float(float_val) => Number::from_f64(*float_val)
                                           .map(Value::Number)
                                           .map(Some)
                                           .ok_or_else(|| Error::msg("Invalid float value")),
            Data::Bool(bool_val) => Ok(Some(Value::Bool(*bool_val))),
            Data::String(str_val) => Ok(Some(Value::String(str_val.to_string()))),
            Data::Error(_) => Err(Error::msg("Error in cell")),
            Data::DateTime(dt) => Ok(Some(Value::String(conversions::excel_datetime(dt.as_f64())?))),
            Data::DurationIso(duration_iso) => Ok(Some(Value::String(duration_iso.to_string()))),
            _ => Err(Error::msg("Unsupported data type"))
        }
    } else {
        Ok(None)
    }
}

pub fn extract_values(sheet: &Range<Data>, instructions: &Map<String, Value>) -> Result<Map<String, Value>, Error> {
    let mut results = Map::new();
    for (key, value) in instructions {
        let (row, col) = if let Some(cell_address) = value.as_str() {
            conversions::address_to_row_col(cell_address)?
        } else if let Some(obj) = value.as_object() {
            let row = obj.get("row").and_then(Value::as_u64).ok_or_else(|| Error::msg("Missing 'row'"))? as u32;
            let col = obj.get("col").and_then(Value::as_u64).ok_or_else(|| Error::msg("Missing 'col'"))? as u32;
            (row, col)
        } else {
            return Err(Error::msg("Invalid or missing row/column specification"));
        };

        match extract_cell_value(sheet, row, col) {
            Ok(Some(cell_value)) => { results.insert(key.clone(), cell_value); },
            Ok(None) => { results.insert(key.clone(), Value::Null); },
            Err(e) => return Err(e),
        }
    }
    Ok(results)
}
