use calamine::{open_workbook_auto, Reader, Range, Data, DataType};
use serde_json::Value;
use anyhow::{Result, Error};
use std::collections::HashMap;

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
pub async fn process_file(file_path: String, extraction_details: Vec<Value>) -> Result<Value, Error> {
    // Initialize an empty array to hold the result for each extraction_detail
    let mut result = Vec::new();
    let mut final_result = serde_json::Map::new();
    final_result.insert("file".to_string(), Value::String(file_path.clone()));

    let mut workbook = open_workbook_auto(&file_path).map_err(Error::new)?;
    // Adjusted to handle different sheet types
    for extract in extraction_details.iter() {
        let mut extraction_result = serde_json::Map::new();
        let map = match extract {
            Value::Object(map) => map,
            _ => return Err(Error::msg("Extraction detail should be a JSON object")),
        };
        let sheet_names = map
            .get("sheets")
            .and_then(|sheets| sheets.as_array())
            .map(|sheet_array| {
                sheet_array.iter().map(|name| {
                    name.as_str().ok_or_else(|| Error::msg("\"sheets\" array should contain only strings"))
                }).collect::<Result<Vec<&str>, Error>>()  // Collecting directly into a Result
            })
            .unwrap_or_else(|| Err(Error::msg("Missing or invalid \"sheets\" key in extraction details")))
            .map_err(|e| e)?;  // Handling the error once, at the end
        
        let cell_pairs: Option<HashMap<String, String>> = match map
            .get("cells")
            .and_then(|cells| cells.as_object())
            .map(|cells_map| {
                cells_map.iter().map(|(key, value)| {
                    // Attempt to convert each value to a string, or return an error if not possible
                    value.as_str()
                        .map(|v| (key.clone(), v.to_string()))
                        .ok_or_else(|| Error::msg(format!("Value for key '{}' is not a string", key)))
                })
                // Collect into a Result<HashMap<String, String>, Error>
                .collect::<Result<HashMap<String, String>, Error>>()
            })
        {
            Some(Ok(map)) => Some(map), // If everything is OK, use the HashMap
            Some(Err(e)) => {
                // Log the error, or handle it as needed
                println!("Error processing cells: {}", e);
                None // Decide to proceed without cell_pairs in case of an error
            },
            None => None, // If "cells" was not present, proceed without cell_pairs
        };
        for sheet_name in &sheet_names {
            // Initialize an object to hold the key-value pairs for this sheet
            let mut sheet_obj = serde_json::Map::new();

            // Load the sheet into memory
            let sheet = workbook
                .worksheet_range(sheet_name)
                .map_err(|e| Error::msg(format!("Failed to access sheet '{}': {}", sheet_name, e)))?;
            // Assuming you're doing something with sheet_names_vec here
        
            if let Some(cell_map) = &cell_pairs {
                for (key, val) in cell_map {
                    if let Some(extracted_value) = extract_cell_value(&sheet, val)? {
                        // Insert the key-value pair into the sheet object
                        sheet_obj.insert(key.clone(), extracted_value);
                    }
                }
            }
    
            // Add the sheet object to the extraction_result object
            extraction_result.insert(sheet_name.to_string(), Value::Object(sheet_obj));
        }
        result.push(Value::Object(extraction_result));
    }

    // Construct the return value as a JSON object containing the array of sheet names
    // Add the data array to the final result
    final_result.insert("data".to_string(), Value::Array(result));

    Ok(Value::Object(final_result)) // Return the constructed JSON object
}
