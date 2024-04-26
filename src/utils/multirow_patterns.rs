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
                for (column_name, column_index_value) in columns {
                    let column_values = match column_index_value {
                        Value::Array(arr) => arr.clone(),
                        Value::String(s) => vec![Value::String(s.clone())],
                        _ => return Err(Error::msg("Invalid column specification")),
                    };
                    if column_values.len() == 1 {
                        let column_index_str = match column_values.first() {
                            Some(Value::String(s)) => s.clone(),
                            _ => return Err(Error::msg("Invalid column specification")),
                        };

                        let col = conversions::column_name_to_index(&column_index_str)?;
                        let value = manipulations::extract_cell_value(sheet, row, col)?;

                        if let Some(value) = value {
                            row_data.insert(column_name.clone(), value);
                        } else {
                            row_data.insert(column_name.clone(), Value::Null);
                        }
                    } else if !column_values.is_empty() {
                        let mut cell_values = Vec::new();
                        for column_index_value in column_values {
                            let column_index_str = match column_index_value {
                                Value::String(s) => s,
                                _ => return Err(Error::msg("Invalid column specification")),
                            };

                            let col = conversions::column_name_to_index(&column_index_str)?;
                            let value = manipulations::extract_cell_value(sheet, row, col)?;

                            if let Some(value) = value {
                                if !value.is_null() {
                                    cell_values.push(value);
                                }
                            }
                        }
                        row_data.insert(column_name.clone(), Value::Array(cell_values));
                    }
                }

                let mut unique_key = unique_id.to_string();
                let mut counter = 1;
                while results.contains_key(&unique_key) {
                    unique_key = format!("{}_{}", unique_id.to_string(), counter);
                    counter += 1;
                }
                results.insert(unique_key, Value::Object(row_data.clone()));
            }
        }
    }
    Ok(results)
}
