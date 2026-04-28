# XLPlainText

Convert an XLSX workbook to plain text using Excel's *xlTextPrinter* save format.

## Description

**XLPlainText** is a Python utility that converts *Excel* XLSX workbooks to plain text files. If the workbook was laid out using a fixed-width font, the conversion will result in a text file with columns aligned as closely as possible to the original layout.

The conversion process:
1. Creates a temporary copy of the XLSX file (so the original can remain open if desired).
2. Uses *Excel* COM automation to save the copy as text using the *xlTextPrinter* format.
3. Cleans up the temporary file.

## Requirements

- **Python**: 3.8 or higher
- **Platform**: *Windows*, *Excel*
- **Dependencies**: `pywin32` >= 306

## Installation

```
pip install xlplaintext
```

## Usage

### Command Line
```
python xlplaintext.py [-h] [-f] xlsx_file [txt_file]
```

### Arguments

- `xlsx_file`: Input XLSX file path (required)
- `txt_file`: Output text file path (optional, defaults to same name with .txt extension)
- `-f, --force`: Force overwrite of the output file if it already exists

### Examples

Convert with automatic output filename (input.xlsx → input.txt):
```
python xlplaintext.py input.xlsx
```

Specify an output filename:
```
python xlplaintext.py input.xlsx output.txt
```

Force overwrite of an existing output file:
```
python xlplaintext.py -f input.xlsx output.txt
python xlplaintext.py --force input.xlsx output.txt
```