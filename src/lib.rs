use pyo3::types::{PyList, PyDict};
use pyo3::prelude::*;
use serde_json::{Value, json};
use std::collections::HashMap;
use tokio::runtime;

mod parallel;
mod read_excel;
use parallel::process_files;

#[pyfunction]
fn excel_extract(_py: Python, file_paths: &PyList, extraction_details: &PyList, num_workers: Option<usize>) -> PyResult<Vec<Vec<String>>> {
    let file_paths: Vec<String> = file_paths.iter().map(|p| {
        p.extract::<String>()
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Error extracting string: {}", e)))
    }).collect::<PyResult<Vec<String>>>()?;

    // Convert PyList extraction_details into Serde Value
    let extraction_details_serde: Vec<Value> = extraction_details.iter().map(|item| {
        let detail_dict = item.downcast::<PyDict>()?.into_iter();
        let mut detail_map = HashMap::new();

        for (k, v) in detail_dict {
            let key: String = k.extract()?;
            let value: Vec<String> = v.extract::<&PyList>()?.iter().map(|item| {
                item.extract::<String>()
                    .map_err(|e| pyo3::exceptions::PyTypeError::new_err(format!("Expected a list of strings, got error: {}", e)))
            }).collect::<PyResult<Vec<String>>>()?;
            detail_map.insert(key, value);
        }

        // Convert each detail_map into Serde Value
        Ok(json!(detail_map))
    }).collect::<PyResult<Vec<Value>>>()?;

    // Creating a new Tokio runtime for blocking call
    let rt = runtime::Runtime::new().unwrap();
    let result = rt.block_on(async {
        // Make sure to adjust the process_files function to accept Vec<Value>
        process_files(file_paths, extraction_details_serde, num_workers.unwrap_or(5)).await
    }).map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Error processing files: {}", e)))?;

    Ok(result)
}

#[pymodule]
fn sheet_excavator(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(excel_extract, m)?)?;
    Ok(())
}
