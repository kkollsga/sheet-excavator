use calamine::{Reader, open_workbook_auto};
use serde_json::{Map, Value};
use anyhow::{Result, Error};
use std::iter::Iterator;
use crate::utils::{conversions, manipulations, single_cells, multirow_patterns, match_sheet_names};

fn extend_unique<T: PartialEq>(vec: &mut Vec<T>, value: T) {
    if !vec.contains(&value) {
        vec.push(value);
    }
}



pub async fn process_file(file_path: String, extraction_details: Vec<Value>) -> Result<Value, Error> {
    let mut results = serde_json::Map::new();
    results.insert("filepath".to_string(), Value::String(file_path.clone()));

    for extract in extraction_details.iter() {
        let map = match extract {
            Value::Object(map) => map,
            _ => return Err(Error::msg("Extraction detail should be a JSON object")),
        };

        let mut workbook = open_workbook_auto(&file_path).map_err(Error::new)?;
        let mut sheet_names: Vec<String> = Vec::new();
        if let Some(sheets) = map.get("sheets") {
            if let Some(sheets_array) = sheets.as_array() {
                let skip_sheets = map.get("skip_sheets")
                    .and_then(|v| v.as_array())
                    .map(|arr| arr.iter().cloned().collect::<Vec<_>>())
                    .unwrap_or_else(|| Vec::new());
                
                for sheet in sheets_array {
                    if let Some(sheet_str) = sheet.as_str() {
                        if sheet_str.contains('*') {
                            for sheet_name in match_sheet_names(&workbook.sheet_names().to_vec(), sheet_str) {
                                if !skip_sheets.iter().any(|s| s == &sheet_name) {
                                    extend_unique(&mut sheet_names, sheet_name);
                                }
                            }
                        } else {
                            let sheet_name = sheet_str.to_string();
                            if !skip_sheets.iter().any(|s| s == &sheet_name) {
                                extend_unique(&mut sheet_names, sheet_name);
                            }
                        }
                    } else {
                        return Err(Error::msg("Invalid sheet name"));
                    }
                }
            } else {
                return Err(Error::msg("Invalid \"sheets\" value in extraction details"));
            }
        } else {
            return Err(Error::msg("Missing \"sheets\" key in extraction details"));
        }
        let break_if_null = map.get("break_if_null").and_then(|f| f.as_str());

        let extractions = map
            .get("extractions")
            .and_then(|extr| extr.as_array())
            .ok_or_else(|| Error::msg("Missing or invalid \"extractions\" key in extraction details"))?
            .iter()
            .map(|extr| {
                let obj = extr.as_object().ok_or_else(|| Error::msg("Each extraction should be a JSON object"))?;
                let function = obj.get("function")
                    .and_then(|f| f.as_str())
                    .ok_or_else(|| Error::msg("Missing 'function' key"))?
                    .to_string();
                let function_label = obj.get("label")
                    .and_then(|f| f.as_str())
                    .unwrap_or("") // Provide "" as default if 'label' is missing
                    .to_string();
                let instructions = obj.get("instructions")
                    .and_then(|i| i.as_object())
                    .cloned()
                    .ok_or_else(|| Error::msg("Missing 'instructions' key"))?;
                Ok((function, function_label, instructions))
            })
            .collect::<Result<Vec<(String, String, Map<String, Value>)>, Error>>()?;

        
        for sheet_name in &sheet_names {
            let sheet = match workbook.worksheet_range(sheet_name) {
                Ok(sheet) => sheet,
                Err(_) => {
                    println!("Warning: Sheet '{}' not found, skipping.", sheet_name);
                    continue;
                }
            };
            if let Some(break_if_null_value) = break_if_null {
                let (row, col) = conversions::address_to_row_col(break_if_null_value)?;
                match manipulations::extract_cell_value(&sheet, row, col) {
                    Ok(Some(cell_value)) => {
                        if cell_value.is_null() {
                            println!("Break condition met at cell {:?}", break_if_null_value);
                            break; // Break out of the sheet loop
                        }
                    },
                    Ok(None) => {
                        println!("Cell {:?} is null", break_if_null_value);
                    },
                    Err(e) => return Err(e),
                }
            }


            let mut sheet_results = Map::new();
            for (function, label, instructions) in &extractions {
                let cells_object = match function.as_str() {
                    "single_cells" => single_cells::extract_values(&sheet, &instructions),
                    "multirow_patterns" => multirow_patterns::extract_rows(&sheet, &instructions),
                    _ => {
                        println!("Unsupported function type '{}'", function);
                        continue;
                    }
                }?;

                if label.is_empty() {
                    // If label is empty, merge cells_object into sheet_results with duplicate handling
                    for (key, value) in cells_object {
                        let mut unique_key = key.clone();
                        let mut counter = 1;
                        while sheet_results.contains_key(&unique_key) {
                            unique_key = format!("{}_{}", key, counter);
                            counter += 1;
                        }
                        sheet_results.insert(unique_key, value);
                    }
                } else {
                    // Merge cells_object into sheet_results under the specified label
                    if let Some(Value::Object(existing_map)) = sheet_results.get_mut(&label.to_string()) {
                        // If the label already exists, merge the new cells_object into the existing object
                        for (key, value) in cells_object {
                            existing_map.insert(key, value); // Update existing keys or add new keys
                        }
                    } else {
                        // If the label does not exist, simply add it
                        sheet_results.insert(label.clone(), Value::Object(cells_object));
                    }
                }
            }

            if let Some(Value::Object(existing_map)) = results.get_mut(&sheet_name.to_string()) {
                // If sheet_name already exists, merge new sheet results into existing map
                for (key, value) in sheet_results {
                    existing_map.insert(key, value);
                }
            } else {
                // Add new sheet results if not present
                results.insert(sheet_name.to_string(), Value::Object(sheet_results));
            }
        }
    }
    Ok(Value::Object(results))
}
