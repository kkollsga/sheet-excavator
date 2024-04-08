use calamine::open_workbook_auto;
use serde_json::{json, Value};
use anyhow::{Result, Error};
use std::collections::HashMap;

pub async fn process_file(file_path: String, extraction_details: Vec<Value>) -> Result<Value, Error> {
    println!("\n\nProcessing file: {}", file_path);
    let mut workbook = open_workbook_auto(&file_path).map_err(Error::new)?;

    for extract in extraction_details.iter() {
        let map = match extract {
            Value::Object(map) => map,
            _ => return Err(Error::msg("Extraction detail should be a JSON object")),
        };
        println!("MAP: {:?}", &map);
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
                eprintln!("Error processing cells: {}", e);
                None // Decide to proceed without cell_pairs in case of an error
            },
            None => None, // If "cells" was not present, proceed without cell_pairs
        };
        for sheet_name in &sheet_names {
            println!("-- Sheet: {}", sheet_name);
            // Assuming you're doing something with sheet_names_vec here
        
            if let Some(cell_map) = &cell_pairs {
                // cell_pairs is defined, iterate over it
                for (key, value) in cell_map {
                    println!("Key: {}, Value: {}", key, value);
                    // Further processing...
                }
            }
        }
    }

    // Construct the return value as a JSON object containing the array of sheet names
    let result = json!({ "sheetNames": ["neei"] });

    Ok(result) // Return the constructed JSON object
}
