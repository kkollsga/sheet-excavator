use anyhow::{Result, Error};
use calamine::{Range, Data};
use serde_json::{Map, Value};
use crate::utils::{conversions, manipulations};

pub fn extract_values(sheet: &Range<Data>, instructions: &Map<String, Value>) -> Result<Map<String, Value>, Error> {
    let mut results = Map::new();
    for (key, value) in instructions {
        let (row, col) = if let Some(cell_address) = value.as_str() {
            conversions::address_to_row_col(cell_address)?
        } else if let Some(obj) = value.as_object() {
            let row = obj.get("row").and_then(Value::as_u64).ok_or_else(|| Error::msg("Missing 'row'"))? as u32;
            let col = obj.get("col").and_then(Value::as_u64).ok_or_else(|| Error::msg("Missing 'col'"))? as u32;
            (row, col)
        } else {
            return Err(Error::msg("Invalid or missing row/column specification"));
        };

        match manipulations::extract_cell_value(sheet, row, col) {
            Ok(Some(cell_value)) => { results.insert(key.clone(), cell_value); },
            Ok(None) => { results.insert(key.clone(), Value::Null); },
            Err(e) => return Err(e),
        }
    }
    Ok(results)
}
