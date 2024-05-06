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
pip install https://github.com/kkollsga/sheet_excavator/blob/main/wheels/sheet_excavator-0.1.27-cp312-none-win_amd64.whl?raw=true
```
*To upgrade an already installed version of Sheet Excavator, use:*
```
pip install --upgrade https://github.com/kkollsga/sheet_excavator/blob/main/wheels/sheet_excavator-0.1.27-cp312-none-win_amd64.whl?raw=true
```

## Sheet Excavator Usage Guide

### Overview
`sheet_excavator` is a Python library designed to assist in extracting data from Excel sheets. This guide provides an overview of how to use the library and its various features.

### Basic Usage
To get started with `sheet_excavator`, you can follow these steps:

```python
import sheet_excavator
import glob

files = glob.glob(r"D:\temp\*")
extraction_details = [...]  # define extraction details (see below)
results = sheet_excavator.excel_extract(files, extraction_details, 10)
```

### Extraction Details
The `extraction_details` parameter is a list of dictionaries that define the extraction rules for each Excel sheet. Each dictionary contains the following keys:

* `sheets`: A list of sheet names to extract data from.
* `extractions`: A list of extraction rules (see below).
* `skip_sheets`: An optional list of sheet names to skip.
* `break_if_null`: An optional column name to break extraction if null.

### Extraction Rules
The `extractions` key in the `extraction_details` dictionary contains a list of extraction rules. There are three types of extraction rules: `single_cells`, `multirow_patterns`, and `dataframe`.

#### Single Cells Extraction
The `single_cells` extraction rule extracts individual cells from the Excel sheet.

*Example:*
```python
{
    "sheets": ["Sheet1"],
    "extractions": [
        {
            "function": "single_cells",
            "label": "single",
            "instructions": {
                "1": "a1",
                "2": "b2",
                "3": "c3",
                "dato": "d4",
                "datotid": "e5"
            }
        }
    ]
}
```
**Instructions:**

* `instructions`: A dictionary where the keys are the column names and the values are the cell references (e.g., "a1", "b2", etc.).

#### Multirow Patterns Extraction
The `multirow_patterns` extraction rule extracts data from multiple rows in the Excel sheet based on a pattern.

*Example:*
```python
{
    "sheets": ["Generell info og kommentarer"],
    "extractions": [
        {
            "function": "multirow_patterns",
            "label": "deposits",
            "instructions": {
                "row_range": [28, 44],
                "unique_id": "B",
                "columns": {
                    "Deposit": "B",
                    "Discovery_well": "C",
                    "Description": "D",
                    "Oil_low": "E",
                    "Oil_base": "F",
                    "Oil_high": "G",
                    "Cond_low": "H",
                    "Cond_base": "I",
                    "Cond_high": "J",
                    "AssGass_low": "K",
                    "AssGass_base": "L",
                    "AssGass_high": "M",
                    "FriGass_low": "N",
                    "FriGass_base": "O",
                    "FriGass_high": "P"
                }
            }
        }
    ]
}
```
**Instructions:**

* `row_range`: A list of two integers defining the row range to extract.
* `unique_id`: The column to use as a unique identifier.
* `columns`: A dictionary where the keys are the column names and the values are the column letters (e.g., "B", "C", etc.).

#### Dataframe Extraction
The dataframe extraction rule extracts data into a Pandas DataFrame.

*Example:*
```python
{
    "sheets": ["Profil_*"],
    "extractions": [
        {
            "function": "dataframe",
            "label": "tabelTest",
            "instructions": {
                "row_range": [39, 64],
                "column_range": ["H", "AA"],
                "header_row": [34, 35]
            }
        }
    ]
}
```
**Instructions:**

* `row_range`: A list of two integers defining the row range to extract.
* `column_range`: A list of column letters to extract.
* `header_row`: A list of row numbers to use as the header.

By following this guide, you should be able to use the `sheet_excavator` library to extract data from your Excel sheets.

## License
Sheet Excavator is released under the MIT License. See the LICENSE file for more details.