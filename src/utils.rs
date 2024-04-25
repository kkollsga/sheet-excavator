// utils.rs

use pyo3::types::{PyList, PyDict, PyString, PyFloat, PyInt};
use pyo3::prelude::*;
use serde_json::{Value, Map};

pub mod single_cells;
pub mod multirow_patterns;
pub mod conversions;
pub mod manipulations;

fn pydict_to_json_value(pydict: &PyDict) -> PyResult<Value> {
    let mut map = Map::new(); // Use serde_json::Map directly

    for (k, v) in pydict {
        let key: String = k.extract()?;
        let value = python_object_to_value(v)?;
        map.insert(key, value); // Insert into serde_json::Map
    }

    Ok(Value::Object(map)) // Create serde_json::Value::Object
}
/// Helper function to convert Python objects to `serde_json::Value`.
fn python_object_to_value(obj: &PyAny) -> PyResult<Value> {
    if let Ok(list) = obj.downcast::<PyList>() {
        let vec: Vec<Value> = list.iter().map(python_object_to_value).collect::<PyResult<_>>()?;
        Ok(Value::from(vec))
    } else if let Ok(dict) = obj.downcast::<PyDict>() {
        pydict_to_json_value(dict)
    } else if let Ok(s) = obj.downcast::<PyString>() {
        Ok(Value::from(s.to_str()?))
    } else if let Ok(num) = obj.downcast::<PyInt>() {
        Ok(Value::from(num.extract::<i64>()?))
    } else if let Ok(num) = obj.downcast::<PyFloat>() {
        Ok(Value::from(num.extract::<f64>()?))
    } else if obj.is_none() {
        Ok(Value::Null)
    } else {
        // Handle other types or raise an error
        Err(PyErr::new::<pyo3::exceptions::PyTypeError, _>(format!("Unsupported Python type: {:?}", obj)))
    }
}

/// Converts a `PyList` of `PyDicts` into a `Vec<serde_json::Value>`.
pub fn pylist_to_json(pylist: &PyList) -> PyResult<Vec<Value>> {
    pylist.iter().map(|item| {
        let detail_dict = item.downcast::<PyDict>()?;
        pydict_to_json_value(detail_dict)
    }).collect()
}


pub fn match_sheet_names(sheet_names: &[String], pattern: &str) -> Vec<String> {
    let (start, end) = match pattern.find('*') {
        Some(index) => {
            let (start, end) = pattern.split_at(index);
            let end = end.trim_start_matches('*');
            (start, end)
        },
        None => ("", ""), // Default start and end if asterisk is not found
    };

    sheet_names
        .iter()
        .filter_map(|sheet_name| {
            if sheet_name.starts_with(start) && sheet_name.ends_with(end) {
                Some(sheet_name.clone())
            } else {
                None
            }
        })
        .collect()
}
