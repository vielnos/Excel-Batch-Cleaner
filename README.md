# Excel Batch Sheet Cleaner

This Python script automates cleaning of Excel files in bulk.

## Features

- Delete top N rows from specified sheets
- Remove shapes/images anchored in deleted rows
- Process multiple files in a folder
- Save cleaned files to a separate output folder
- Works with `.xlsx` and `.xls` files
- Fully interactive sheet and row configuration

## Requirements

- Windows
- Python 3.x
- pywin32 package (`pip install pywin32`)

## Usage

```bash
python clean_excel.py
