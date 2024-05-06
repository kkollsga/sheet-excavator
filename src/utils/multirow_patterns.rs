use anyhow::{Result, Error};
use calamine::{Range, Data};
use serde_json::{Map, Value};
use crate::utils::{conversions, manipulations};

pub fn extract_rows(sheet: &Range<Data>, instructions: &Map<String, Value>) -> Result<Map<String, Value>, Error> {
    let mut results = Map::new();

    // Retrieve and parse the row_range array
    let row_range = instructions.get("row_range")
        .and_then(Value::as_array)
        .ok_or_else(|| Error::msg("Missing or invalid 'row_range'"))?;
    let start_row = row_range.get(0).and_then(Value::as_u64)
        .ok_or_else(|| Error::msg("Missing 'start_row' in 'row_range'"))? as u32;  // Adjust for zero-based index
    let end_row = row_range.get(1).and_then(Value::as_u64)
        .ok_or_else(|| Error::msg("Missing 'end_row' in 'row_range'"))? as u32;  // Adjust for zero-based index

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
        match manipulations::extract_cell_value(sheet, row, unique_id_index, false) {
            Ok((Some(unique_id), _)) if unique_id != Value::Null => {
                for (column_name, column_index_value) in columns {
                    let column_values = match column_index_value {
                        Value::Array(arr) => arr.clone(),
                        Value::String(s) => vec![Value::String(s.clone())],
                        _ => return Err(Error::msg("Invalid column specification")),
                    };

                    let mut cell_values = Vec::new();
                    for column_index_value in column_values {
                        let column_index_str = match column_index_value {
                            Value::String(s) => s,
                            _ => return Err(Error::msg("Invalid column specification")),
                        };

                        let col = conversions::column_name_to_index(&column_index_str)?;
                        match manipulations::extract_cell_value(sheet, row, col, false) {
                            Ok((Some(value), _)) if !value.is_null() => cell_values.push(value),
                            Ok((Some(_), _)) => (),  // Handle the case for non-null values that are not needed
                            Ok((None, _)) => (),     // Ignore when no value is found
                            Err(e) => return Err(e), // Propagate errors
                        }
                    }

                    let final_value = match cell_values.len() {
                        0 => Value::Null,
                        1 => cell_values.pop().unwrap(),
                        _ => Value::Array(cell_values),
                    };
                    row_data.insert(column_name.clone(), final_value);
                }

                let mut unique_key = unique_id.to_string();
                let mut counter = 1;
                while results.contains_key(&unique_key) {
                    unique_key = format!("{}_{}", unique_id.to_string(), counter);
                    counter += 1;
                }
                results.insert(unique_key, Value::Object(row_data));
            },
            Ok(_) => (), // Ignore null or None unique_ids
            Err(e) => return Err(e),
        }
    }
    Ok(results)
}
