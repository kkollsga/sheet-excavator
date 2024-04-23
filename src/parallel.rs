use tokio::sync::Semaphore; // Import Semaphore from tokio::sync
use std::{sync::Arc, path::Path, time::Instant}; // Import Arc for creating reference-counted pointers
use futures::stream::{FuturesUnordered, StreamExt}; // Import FuturesUnordered and StreamExt for managing and polling futures
use serde_json::{Value, Map}; // Import serde_json::Value
use crate::read_excel::process_file;
use anyhow::{Result, Error};// Use anyhow::Result for simplified error handling


// Helper function to extract the base filename without extension
fn extract_filename(path: &str) -> String {
    Path::new(path)
        .file_stem()
        .and_then(|name| name.to_str())
        .unwrap_or_default()
        .to_string()
}

pub async fn process_files(file_paths: Vec<String>, extraction_details: Vec<Value>, num_workers: usize) -> Result<Map<String, Value>, Error> {
    println!("Processing files!");
    let semaphore = Arc::new(Semaphore::new(num_workers));

    let mut futures = FuturesUnordered::new();
    let start_time = Instant::now();
    let total = file_paths.len();

    for (index, path_str) in file_paths.into_iter().enumerate() {
        let path_str_clone = path_str.clone();
        let details_clone = extraction_details.clone();
        let sem_clone = semaphore.clone();

        let permit = sem_clone.acquire_owned().await.unwrap();

        futures.push(tokio::spawn(async move {
            let result = process_file(path_str_clone, details_clone).await;
            let files_left = total - index + 1;
            let avg_time_per_file = if index > 0 {
                start_time.elapsed().as_secs_f64() / index as f64
            } else {
                0.0
            };
            let estimated_time_left = avg_time_per_file * files_left as f64;
            println!("Progress: {}/{} files. Avg: {:.2}s. Time left: {:.2}s.", index, total, avg_time_per_file, estimated_time_left);
            drop(permit);
            result
        }));
    }

    let mut results = Map::new();
    while let Some(res) = futures.next().await {
        match res {
            Ok(Ok(value)) => {
                if let Some(file_path) = value.get("filepath").and_then(|v| v.as_str()) {
                    let base_filename = extract_filename(file_path);
                    let mut filename_key = base_filename.clone();
                    let mut counter = 1;
                    // Ensure the key is unique by appending a counter if needed
                    while results.contains_key(&filename_key) {
                        filename_key = format!("{}_{}", base_filename, counter);
                        counter += 1;
                    }
                    results.insert(filename_key, value);
                }
            },
            Ok(Err(e)) => return Err(e.into()), // Convert the inner error to the function's error type
            Err(e) => return Err(anyhow::Error::new(e)), // Convert the JoinError to the function's error type
        }
    }

    println!("All files processed. Total time: {:.2?}", start_time.elapsed());
    Ok(results)
}