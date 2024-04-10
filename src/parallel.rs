use tokio::sync::Semaphore; // Import Semaphore from tokio::sync
use std::sync::Arc; // Import Arc for creating reference-counted pointers
use futures::stream::{FuturesUnordered, StreamExt}; // Import FuturesUnordered and StreamExt for managing and polling futures
use serde_json::Value; // Import serde_json::Value
use crate::read_excel::process_file;
use anyhow::Result; // Use anyhow::Result for simplified error handling
use std::time::Instant; // Import Instant to measure elapsed time

pub async fn process_files(file_paths: Vec<String>, extraction_details: Vec<Value>, num_workers: usize) -> Result<Vec<Value>> {
    println!("Processing files!");
    let semaphore = Arc::new(Semaphore::new(num_workers)); // Wrap Semaphore in an Arc for shared ownership

    let mut futures = FuturesUnordered::new(); // Create a FuturesUnordered collection for managing futures
    let start_time = Instant::now(); // Record the start time for logging progress

    for (_index, path_str) in file_paths.into_iter().enumerate() {
        let path_str_clone = path_str.clone();
        let details_clone = extraction_details.clone(); // Clone extraction_details for each async task
        let sem_clone = semaphore.clone(); // Clone the Arc, not the Semaphore itself

        let permit = sem_clone.acquire_owned().await.unwrap(); // Acquire a permit from the semaphore

        futures.push(tokio::spawn(async move {
            // Once a permit is acquired, push the task into FuturesUnordered
            println!("Pushing. {}",&path_str);
            let result = process_file(path_str_clone, details_clone).await;
            println!("Done. {}",&path_str);
            drop(permit); // Release the permit when the task is done
            result
        }));
    }

    let total_files = futures.len();
    let mut results = Vec::with_capacity(total_files);

    while let Some(res) = futures.next().await {
        // Push the successful results into the results vector
        match res {
            Ok(Ok(value)) => {
                println!("Success.");
                results.push(value); // Handle the double Result layer (tokio::spawn + process_file)
                log_progress(&results, total_files, &start_time); // Log progress after each file is processed
            },
            Ok(Err(e)) => return Err(e.into()), // Convert the inner error to the function's error type
            Err(e) => return Err(anyhow::Error::new(e)), // Convert the JoinError to the function's error type
        }
    }

    println!("All files processed. Total time: {:.2?}", start_time.elapsed()); // Log the completion of all tasks
    Ok(results) // Return the results if successful
}

// Function to log the progress of file processing
fn log_progress(results: &[Value], total_files: usize, start_time: &Instant) {
    let processed_files = results.len();
    let files_left = total_files - processed_files;
    let avg_time_per_file = if processed_files > 0 {
        start_time.elapsed().as_secs_f64() / processed_files as f64
    } else {
        0.0 // Avoid division by zero if no files have been processed yet
    };
    let estimated_time_left = avg_time_per_file * files_left as f64;
    println!(
        " Progress: {}/{} files. Avg: {:.2}s. Time left: {:.2}s.",
        processed_files, total_files, estimated_time_left, avg_time_per_file
    );
}

