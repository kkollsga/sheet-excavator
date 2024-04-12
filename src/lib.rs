use pyo3::types::PyList;
use pyo3::prelude::*;
use serde_json::to_string;
use tokio::runtime;

mod parallel;
mod read_excel;
use parallel::process_files;
mod utils; // Import the utils module
use utils::pylist_to_json; // Import the conversion function

#[pyfunction]
fn excel_extract(_py: Python, file_paths: &PyList, extraction_details: &PyList, num_workers: Option<usize>) -> PyResult<Vec<String>> {
    // Extract file paths from Python list to Rust Vec<String>
    let file_paths: Vec<String> = file_paths.iter().map(|p| {
        p.extract::<String>()
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Error extracting string: {}", e)))
    }).collect::<PyResult<Vec<String>>>()?;
    let extraction_details_serde = pylist_to_json(extraction_details)?;

    // Create a new Tokio runtime to run async process_files function
    let rt = runtime::Runtime::new().unwrap();
    let results = rt.block_on(async {
        let (results, mut progress_receiver) = process_files(file_paths, extraction_details_serde, num_workers.unwrap_or(5)).await?;
        while let Some(progress) = progress_receiver.recv().await {
            println!("Progress: {}", progress);
        }
        Ok(results)
    }).map_err(|e: anyhow::Error| pyo3::exceptions::PyRuntimeError::new_err(format!("Error processing files: {}", e)))?;


    // Convert the Serde JSON Value results into JSON strings
    let json_strings = results.into_iter().map(|val| {
        to_string(&val).map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Error converting to JSON string: {}", e)))
    }).collect::<PyResult<Vec<String>>>()?;
    Ok(json_strings)
}

#[pymodule]
fn sheet_excavator(_py: Python, m: &PyModule) -> PyResult<()> {
    // Register the excel_extract function in the Python module
    m.add_function(wrap_pyfunction!(excel_extract, m)?)?;
    Ok(())
}
