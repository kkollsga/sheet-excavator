use anyhow::{Result, Error};
use calamine::{Range, Data};
use serde_json::{Map, Value};
use crate::utils::{conversions, manipulations};

pub fn extract_rows(sheet: &Range<Data>, instructions: &Map<String, Value>) -> Result<Map<String, Value>, Error> {
    let mut results = Map::new();

    let start_row = instructions
        .get("start_row")
        .and_then(Value::as_u64)
        .ok_or_else(|| Error::msg("Missing 'start_row'"))
        .map(|r| r as u32 - 1)?;  // Subtract 1 to adjust for zero-based index

    let end_row = instructions
        .get("end_row")
        .and_then(Value::as_u64)
        .ok_or_else(|| Error::msg("Missing 'end_row'"))
        .map(|r| r as u32 - 1)?;  // Subtract 1 to adjust for zero-based index

    let columns = instructions
        .get("columns")
        .and_then(Value::as_object)
        .ok_or_else(|| Error::msg("Missing 'columns'"))?;

    let unique_id_column = instructions
        .get("unique_id")
        .and_then(Value::as_str)
        .ok_or_else(|| Error::msg("Missing 'unique_id'"))?;
    let unique_id_index = conversions::column_name_to_index(unique_id_column)?;

    for row in start_row..=end_row {
        let mut row_data = Map::new();
        let unique_id = manipulations::extract_cell_value(sheet, row, unique_id_index)?;
        if let Some(unique_id) = unique_id {
            if unique_id != Value::Null {
                for (column_index, column_name) in columns {
                    let col = conversions::column_name_to_index(column_index)?;
                    let column_name_str = column_name.as_str().unwrap_or_default();
                    let value = manipulations::extract_cell_value(sheet, row, col)?;
        
                    if let Some(value) = value {
                        row_data.insert(column_name_str.to_string(), value);
                    } else {
                        row_data.insert(column_name_str.to_string(), Value::Null);
                    }
                }
                results.insert(unique_id.to_string(), Value::Object(row_data.clone()));
            }
        }
    }
    Ok(results)
}