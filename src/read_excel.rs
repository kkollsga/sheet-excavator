use calamine::{Reader, open_workbook_auto};
use serde_json::{Map, Value};
use anyhow::{Result, Error};

use crate::utils::single_cells;

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
                let instructions = obj.get("instructions")
                    .and_then(|i| i.as_object())
                    .cloned()
                    .ok_or_else(|| Error::msg("Missing 'instructions' key"))?;
                Ok((function, instructions))
            })
            .collect::<Result<Vec<(String, Map<String, Value>)>, Error>>()?;

        let mut workbook = open_workbook_auto(&file_path).map_err(Error::new)?;

        let mut extraction_result = Map::new();
        for sheet_name in &sheet_names {
            let sheet = match workbook.worksheet_range(sheet_name) {
                Ok(sheet) => sheet,
                Err(_) => {
                    println!("Warning: Sheet '{}' not found, skipping.", sheet_name);
                    continue;
                }
            };

            for (function, instructions) in &extractions {
                let cells_object = match function.as_str() {
                    "single_cells" => single_cells::extract_values(&sheet, &instructions),
                    "multiple_cells" => {
                        println!("multicell: {:?}", instructions);
                        continue;  // Placeholder for actual logic
                    },
                    _ => {
                        println!("Unsupported function type '{}'", function);
                        continue;
                    }
                }?;

                extraction_result.insert(sheet_name.to_string(), Value::Object(cells_object));
            }
        }

        if !extraction_result.is_empty() {
            if let Some(Value::Array(data)) = results.get_mut("data") {
                data.push(Value::Object(extraction_result));
            }
        }
    }

    Ok(Value::Object(results))
}
