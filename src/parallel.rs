use futures::future::try_join_all;
use serde_json::Value; // Import serde_json::Value
use crate::read_excel::process_file;
use anyhow::Result; // Use anyhow::Result for simplified error handling

// Update function signature to accept Vec<Value> for extraction_details
pub async fn process_files(file_paths: Vec<String>, extraction_details: Vec<Value>) -> Result<Vec<Vec<String>>> {
    let futures = file_paths.into_iter().map(|path_str| {
        let details_clone = extraction_details.clone(); // Clone extraction_details for each async task
        async move {
            // Call process_file for each file, passing cloned extraction details
            process_file(path_str, details_clone).await
        }
    });

    // Use try_join_all to await all futures and collect results
    let results = try_join_all(futures).await?;
    println!("\n\nRust results: {:?}", &results);

    Ok(results) // Return the results if successful
}
