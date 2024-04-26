# Sheet Excavator


## Overview
Sheet Excavator is a Rust-based tool designed to facilitate the efficient extraction of data from standardized Excel forms. Traditional reporting often relies on Excel forms that do not conform to the typical CSV data storage format, making data extraction challenging. Existing Python-based workflows may also suffer from performance issues when handling large databases of forms stored in .xlsx files.

Leveraging Rust's high performance and robust multithreading capabilities, Sheet Excavator provides a powerful API tailored for extracting data from unstructured Excel layouts. It supports various functionalities including single cell extraction, row-based patterns, and multi-column arrays, returning results in an easy-to-use JSON format.

## Key features
- High Performance: Utilizes Rust’s efficiency and multithreading to handle large datasets.
- Flexible Data Extraction: Supports various extraction methods for complex Excel form layouts.
- JSON Output: Seamlessly integrates with modern data pipelines by outputting data in JSON format.

## Installation
Sheet Excavator is currently in its alpha phase, with limited functionality and is only compatible with Python 3.12 on Windows (AMD64 architecture). You can install the library directly from the repository using pip.

### Prerequisites
Python 3.12
Windows AMD64 system

### Install with pip
*To install Sheet Excavator, run the following command in your terminal:*
```
pip install https://github.com/kkollsga/sheet_excavator/blob/main/wheels/sheet_excavator-0.1.22-cp312-none-win_amd64.whl?raw=true
```
*To upgrade an already installed version of Sheet Excavator, use:*
```
pip install --upgrade https://github.com/kkollsga/sheet_excavator/blob/main/wheels/sheet_excavator-0.1.22-cp312-none-win_amd64.whl?raw=true
```

## Usage
Detailed documentation on how to use Sheet Excavator will be provided here or linked to a documentation site.

## License
Sheet Excavator is released under the MIT License. See the LICENSE file for more details.