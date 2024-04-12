use pyo3::types::PyList;
use pyo3::prelude::*;
use pyo3::types::PyAny;
use serde_json::to_string;
use tokio::runtime;
use pyo3_asyncio::tokio::future_into_py;
use anyhow::Error as AnyhowError;
mod parallel;
mod read_excel;
use parallel::process_files;
mod utils; // Import the utils module
use utils::pylist_to_json; // Import the conversion function

#[pyfunction]
fn excel_extract(py: Python, callback: PyObject, file_paths: &PyList, extraction_details: &PyList, num_workers: Option<usize>) -> PyResult<()> {
    let file_paths: Vec<String> = file_paths.iter().map(|p| {
        p.extract::<String>()
            .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Error extracting string: {}", e)))
    }).collect::<PyResult<Vec<String>>>()?;
    
    let extraction_details_serde = pylist_to_json(extraction_details)?;

    // Use the existing runtime or create a new one
    pyo3_asyncio::tokio::get_runtime().block_on(async move {
        let (results, mut progress_receiver) = process_files(file_paths, extraction_details_serde, num_workers.unwrap_or(5)).await
    .map_err(|e| pyo3::exceptions::PyRuntimeError::new_err(format!("Error processing files: {}", e)))?;

        while let Some(progress) = progress_receiver.recv().await {
            // Call the provided Python callback function asynchronously with the progress
            let _ = Python::with_gil(|py| {
                let _call_result = callback.call1(py, (progress,));
            });
        }

        // Send the final results as a JSON string to the callback
        let json_results: Vec<String> = results.into_iter().map(|val| {
            to_string(&val).unwrap() // Assuming no errors for simplicity
        }).collect();

        Python::with_gil(|py| {
            let _call_result = callback.call1(py, ("done", json_results));
        });

        Ok(())
    })
}

#[pymodule]
fn sheet_excavator(_py: Python, m: &PyModule) -> PyResult<()> {
    // Register the excel_extract function in the Python module
    m.add_function(wrap_pyfunction!(excel_extract, m)?)?;
    Ok(())
}
