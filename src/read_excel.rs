use calamine::{Reader, open_workbook_auto};
use serde_json::{Map, Value};
use anyhow::{Result, Error};

use crate::utils::{single_cells, multirow_patterns};

pub async fn process_file(file_path: String, extraction_details: Vec<Value>) -> Result<Value, Error> {
    let mut results = serde_json::Map::new();
    results.insert("file".to_string(), Value::String(file_path.clone()));
    results.insert("data".to_string(), Value::Array(Vec::new()));  // Initialize "data" as an empty array

    for extract in extraction_details.iter() {
        let map = match extract {
            Value::Object(map) => map,
            _ => return Err(Error::msg("Extraction detail should be a JSON object")),
        };

        let sheet_names = map
            .get("sheets")
            .and_then(|sheets| sheets.as_array())
            .ok_or_else(|| Error::msg("Missing or invalid \"sheets\" key in extraction details"))?
            .iter()
            .map(|name| name.as_str().ok_or_else(|| Error::msg("\"sheets\" array should contain only strings")))
            .collect::<Result<Vec<&str>, Error>>()?;

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

        let mut workbook = open_workbook_auto(&file_path).map_err(Error::new)?;
        for sheet_name in &sheet_names {
            let sheet = match workbook.worksheet_range(sheet_name) {
                Ok(sheet) => sheet,
                Err(_) => {
                    println!("Warning: Sheet '{}' not found, skipping.", sheet_name);
                    continue;
                }
            };
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
                    // Insert the cells_object under the key specified by label
                    // Check if the key already exists and append suffix if needed
                    let mut unique_label = label.clone();
                    let mut counter = 1;
                    while sheet_results.contains_key(&unique_label) {
                        unique_label = format!("{}_{}", label, counter);
                        counter += 1;
                    }
                    sheet_results.insert(unique_label, Value::Object(cells_object));
                }
            }

            if !sheet_results.is_empty() {
                // Wrap sheet results in an object with the sheet name as the key
                let mut sheet_data = Map::new();
                sheet_data.insert(sheet_name.to_string(), Value::Object(sheet_results));

                // Append to the data array under results
                if let Some(Value::Array(data)) = results.get_mut("data") {
                    data.push(Value::Object(sheet_data));
                }
            }
        }

    }
    Ok(Value::Object(results))
}
