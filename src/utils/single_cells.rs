use anyhow::{Result, Error};
use calamine::{Range, Data, DataType};
use serde_json::{Map, Value};

fn cell_address_to_row_col(cell_address: &str) -> Result<(u32, u32), Error> {
    let split_at = cell_address.chars().position(|c| c.is_digit(10)).ok_or_else(|| Error::msg("Invalid cell address format"))?;
    let (col_str, row_str) = cell_address.split_at(split_at);

    let col = col_str.chars().rev().enumerate().try_fold(0u32, |acc, (i, c)| {
        if let Some(digit) = c.to_digit(36) {
            Ok(acc + (digit - 9) * 26u32.pow(i as u32))
        } else {
            Err(Error::msg("Invalid column label"))
        }
    })?;

    let row: u32 = row_str.parse().map_err(|_| Error::msg("Invalid row number"))?;

    // Adjust for 0-based indexing used by Calamine
    Ok((row - 1, col - 1))
}
fn extract_cell_value(sheet: &Range<Data>, cell_address: &str) -> Result<Option<Value>, Error> {
    let (row, col) = cell_address_to_row_col(cell_address)?;
    if let Some(cell) = sheet.get_value((row, col)) {
        if cell.is_empty() {
            Ok(Some(Value::Null))
        } else if let Some(int_val) = cell.get_int() {
            Ok(Some(Value::Number(int_val.into())))
        } else if let Some(float_val) = cell.get_float() {
            Ok(Some(Value::Number(serde_json::Number::from_f64(float_val).expect("Invalid float value"))))
        } else if let Some(bool_val) = cell.get_bool() {
            Ok(Some(Value::Bool(bool_val)))
        } else if let Some(str_val) = cell.get_string() {
            Ok(Some(Value::String(str_val.to_string())))
        } else {
            Err(Error::msg("Unrecognized cell type"))
        }
    } else {
        Ok(None)
    }
}

pub fn extract_values(sheet: &Range<Data>, instructions: &Map<String, Value>) -> Result<Map<String, Value>, Error> {
    let mut results = Map::new();
    for (key, value) in instructions {
        if let Some(cell_address) = value.as_str() {
            match extract_cell_value(sheet, cell_address) {
                Ok(Some(cell_value)) => { results.insert(key.clone(), cell_value); },
                Ok(None) => { results.insert(key.clone(), Value::Null); },
                Err(e) => return Err(e),
            }
        }
    }
    Ok(results)
}