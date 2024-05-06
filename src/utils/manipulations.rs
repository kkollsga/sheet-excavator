use anyhow::{Result, Error};
use calamine::{Range, Data};
use serde_json::{Value, Number};
use crate::utils::conversions;

pub fn extract_cell_value(sheet: &Range<Data>, row: u32, col: u32, force_str: bool) -> Result<(Option<Value>, String), Error> {
    let cell = sheet.get_value((row-1, col));
    if cell.is_none() {
        return Ok((None, "Null".to_string()));
    }
    let cell = cell.unwrap(); // Safe to unwrap here because we know it's Some

    if force_str {
        // When forcing to string, handle all types as string output
        let description = match &cell {
            Data::Empty => "Null".to_string(),
            Data::Int(int_val) => int_val.to_string(),
            Data::Float(float_val) => float_val.to_string(),
            Data::Bool(bool_val) => bool_val.to_string(),
            Data::String(str_val) => str_val.trim().to_string(),
            Data::Error(_) => return Err(Error::msg("Error in cell")),
            Data::DateTime(dt) => conversions::excel_datetime(dt.as_f64())?,
            Data::DurationIso(duration_iso) => duration_iso.to_string(),
            _ => return Err(Error::msg("Unsupported data type")),
        };
        let value = Value::String(description.clone());
        return Ok((Some(value), description));
    }
    let result = match cell {
        Data::Empty => (Some(Value::Null), "Null".to_string()),
        Data::Int(int_val) => (Some(Value::Number(Number::from(*int_val))), "Int".to_string()),
        Data::Float(float_val) => match Number::from_f64(*float_val) {
            Some(number) => (Some(Value::Number(number)), "Float".to_string()),
            None => return Err(Error::msg("Invalid float value")),
        },
        Data::Bool(bool_val) => (Some(Value::Bool(*bool_val)), "Bool".to_string()),
        Data::String(str_val) => {
            let trimmed_str = str_val.trim().to_string();
            (Some(Value::String(trimmed_str)), "String".to_string())
        },
        Data::Error(_) => return Err(Error::msg("Error in cell")),
        Data::DateTime(dt) => (
            Some(Value::String(conversions::excel_datetime(dt.as_f64())?)),
            "DateTime".to_string(),
        ),
        Data::DurationIso(duration_iso) => (
            Some(Value::String(duration_iso.to_string())),
            "DurationIso".to_string(),
        ),
        _ => return Err(Error::msg("Unsupported data type")),
    };
    Ok(result)
}

pub fn extract_headers(
    sheet: &Range<Data>,
    header_rows: &[u32], // Array of row indices
    column: u32,         // Column index
    separator: &str      // Separator to join headers
) -> Result<String, Error> {
    let mut headers = Vec::new();

    for &row in header_rows {
        let (_, header_description) = extract_cell_value(sheet, row, column, true)?; // Force string extraction
        // Replace carriage returns and new lines with a space
        let clean_header = header_description.replace("\r\n", " ").replace("\n", " ").replace("\r", " ");
        headers.push(clean_header);
    }

    // Join the cleaned headers with the specified separator
    Ok(headers.join(separator))
}

pub fn extract_column_data(
    sheet: &Range<Data>,
    column: u32,
    start_row: u32,
    end_row: u32
) -> Result<Value, Error> {
    let mut column_data: Vec<Value> = Vec::new();

    for row in start_row..=end_row {
        let (cell_value, _) = extract_cell_value(sheet, row, column, false)?;
        // If cell_value is None, insert null into the column_data
        let value = match cell_value {
            Some(v) => v,
            None => Value::Null,
        };
        column_data.push(value);
    }

    Ok(Value::Array(column_data))
}