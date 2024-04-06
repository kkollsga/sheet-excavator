use calamine::open_workbook_auto;
use serde_json::Value; // Import serde_json::Value
use anyhow::{Result, Error};

// Update the function signature to accept extraction_details as Vec<Value>
pub async fn process_file(file_path: String, extraction_details: Vec<Value>) -> Result<Vec<String>, Error> {
    println!("\n\nProcessing file: {}", file_path);
    let mut workbook = match open_workbook_auto(&file_path) {
        Ok(wb) => wb,
        Err(e) => return Err(Error::new(e)),
    };

    for extract in extraction_details.iter() {
        // Assume each extract is an object and use match to navigate
        match extract {
            Value::Object(map) => {
                // Use match to check for "sheets" key in the map
                match map.get("sheets") {
                    Some(sheets) => {
                        // Use match to ensure sheets is an array
                        match sheets.as_array() {
                            Some(sheet_names) => {
                                for sheet_name in sheet_names {
                                    // Use match to ensure each sheet_name is a string
                                    match sheet_name.as_str() {
                                        Some(name) => println!("-- Sheet: {}", name),
                                        None => return Err(Error::msg("Sheet name should be a string")),
                                    }
                                }
                            },
                            None => return Err(Error::msg("\"sheets\" key should map to an array of strings")),
                        }
                    },
                    None => return Err(Error::msg("Missing \"sheets\" key in extraction details")),
                }
            },
            _ => return Err(Error::msg("Extraction detail should be a JSON object")),
        }
    }

    Ok(vec![]) // Return an empty Vec for now, replace with actual processing results
}
