use tokio::sync::Semaphore; // Import Semaphore from tokio::sync
use std::sync::Arc; // Import Arc for creating reference-counted pointers
use futures::stream::{FuturesUnordered, StreamExt}; // Import FuturesUnordered and StreamExt for managing and polling futures
use serde_json::Value; // Import serde_json::Value
use crate::read_excel::process_file;
use anyhow::Result; // Use anyhow::Result for simplified error handling
use std::time::Instant; // Import Instant for timing

pub async fn process_files(file_paths: Vec<String>, extraction_details: Vec<Value>, num_workers: usize) -> Result<Vec<Value>> {
    let semaphore = Arc::new(Semaphore::new(num_workers));

    let mut futures = FuturesUnordered::new();
    let total_files = file_paths.len();
    let mut processed_files = 0;
    let start_time = Instant::now();

    for path_str in file_paths.iter().cloned() {
        let details_clone = extraction_details.clone();
        let sem_clone = semaphore.clone();
        // Acquire a permit before spawning the task
        let _permit = sem_clone.acquire_owned().await.expect("Failed to acquire semaphore permit");
        futures.push(tokio::spawn(async move {
            let result = process_file(path_str.clone(), details_clone).await;
            result
        }));
    }

    let mut results = Vec::new();

    while let Some(res) = futures.next().await {
        match res {
            Ok(Ok(value)) => {
                processed_files += 1;
                let avg_time_per_file = start_time.elapsed().as_secs_f64() / processed_files as f64;
                let files_left = total_files - processed_files;
                let estimated_time_left = avg_time_per_file * files_left as f64;

                // Format the average time and estimated time left with two decimal places
                let avg_time_str = format!("{:.2}", avg_time_per_file);
                let estimated_time_str = format!("{:.2}", estimated_time_left);

                println!("Progress: {}/{} files processed. Average time per file: {} seconds. Estimated time left: {} seconds", processed_files, total_files, avg_time_str, estimated_time_str);
                results.push(value);
            },
            Ok(Err(e)) => return Err(e.into()),
            Err(e) => return Err(anyhow::Error::new(e)),
        }
    }

    Ok(results)
}
