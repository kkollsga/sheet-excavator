use tokio::sync::Semaphore; // Import Semaphore from tokio::sync
use std::sync::Arc; // Import Arc for creating reference-counted pointers
use futures::stream::{FuturesUnordered, StreamExt}; // Import FuturesUnordered and StreamExt for managing and polling futures
use serde_json::Value; // Import serde_json::Value
use crate::read_excel::process_file;
use anyhow::Result; // Use anyhow::Result for simplified error handling

pub async fn process_files(file_paths: Vec<String>, extraction_details: Vec<Value>, num_workers: usize) -> Result<Vec<Value>> {
    let semaphore = Arc::new(Semaphore::new(num_workers)); // Wrap Semaphore in an Arc for shared ownership

    let mut futures = FuturesUnordered::new(); // Create a FuturesUnordered collection for managing futures

    for path_str in file_paths {
        let details_clone = extraction_details.clone(); // Clone extraction_details for each async task
        let sem_clone = semaphore.clone(); // Clone the Arc, not the Semaphore itself

        let permit = sem_clone.acquire_owned().await.unwrap(); // Acquire a permit from the semaphore

        futures.push(tokio::spawn(async move {
            // Once a permit is acquired, push the task into FuturesUnordered
            let result = process_file(path_str, details_clone).await;
            drop(permit); // Release the permit when the task is done
            result
        }));
    }

    let mut results = Vec::new();

    while let Some(res) = futures.next().await {
        // Push the successful results into the results vector
        match res {
            Ok(Ok(value)) => results.push(value), // Handle the double Result layer (tokio::spawn + process_file)
            Ok(Err(e)) => return Err(e.into()), // Convert the inner error to the function's error type
            Err(e) => return Err(anyhow::Error::new(e)), // Convert the JoinError to the function's error type
        }
    }

    Ok(results) // Return the results if successful
}
