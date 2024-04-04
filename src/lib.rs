use pyo3::prelude::*;
use pyo3::types::PyList;

/// A Python module implemented in Rust.
#[pymodule]
fn sheet_excavator(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(excel_extract, m)?)?;
    Ok(())
}

/// Function to test Excel reading functionality.
#[pyfunction]
fn excel_extract(file_path: String, extraction_details: &PyList) {
    println!("Files: {}", file_path);
    for detail in extraction_details.iter() {
        println!("details: {}", detail);
    }
}