# Excel Comparator Tool

A comprehensive Python tool for comparing two Excel files cell-by-cell and identifying all differences with detailed context information.

## Features

- **Cell-by-Cell Comparison**: Compares every cell in all sheets from two Excel files
- **Multi-Format Support**: Works with `.xlsx`, `.xls`, and `.xlsm` files
- **Error Detection**: Identifies Excel error values (#DIV/0!, #N/A, #NAME?, #NULL!, #NUM!, #REF!, #VALUE!)
- **Context Information**: 
  - Shows column headers (row 1) for each difference
  - Includes column D values for reference
  - Displays sheet name and cell location
- **Flexible Export**: Export results as either CSV or Excel (.xlsx) file
- **Sheet Comparison**: Reports missing sheets between files
- **Interactive Input**: User-friendly prompts for file selection with path validation
- **Formatted Output**: Color-coded console output and professionally styled Excel export

## Requirements

- Python 3.7 or higher
- openpyxl (for reading/writing Excel files)

See `requirements.txt` for complete dependencies.

## Installation

1. Clone or download the project
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Running the Tool

```bash
python excel_comparator.py
```

### Workflow

1. **File Selection**: The tool will prompt you to enter the paths for both Excel files
   - Path can be absolute or relative
   - File validation ensures both files exist and are valid Excel files
   
2. **Comparison**: The tool compares all sheets and cells
   - Displays found differences in the console
   - Groups results by sheet name and sorts by row/column
   
3. **Export Option**: After comparison, choose to export results
   - Select format: CSV or Excel
   - Optionally specify output filename

### Example

```
====================================================================================================
Excel File Comparator Tool
====================================================================================================

Enter path for First Excel file: C:\files\report1.xlsx
✓ File loaded: C:\files\report1.xlsx

Enter path for Second Excel file: C:\files\report2.xlsx
✓ File loaded: C:\files\report2.xlsx

Comparing 'report1.xlsx' and 'report2.xlsx'...

Found 5 difference(s):

====================================================================================================

Sheet: 'Data' (5 difference(s))
----------------------------------------------------------------------------------------------------
  Cell: B2
    File 1: 100
    File 2: 105

  Cell: D5
    File 1: 'Complete'
    File 2: 'Pending'
...
```

## CSV Export

The CSV export includes:
- **Sheet**: Name of the sheet containing the difference
- **Cell**: Cell address (e.g., A1, B5)
- **Error_name_1/2**: Column D values from the difference row (for reference)
- **Column**: Column letter
- **Column Header (File 1/2)**: Header value from row 1 of that column
- **File 1/2 Value**: The actual differing values

## Excel Export

The Excel export includes:
- Same columns as CSV export
- Formatted header row (blue background, white text)
- Auto-adjusted column widths
- Professional styling for easy reading

## Output Structure

When differences are found, they are organized by:
1. **Sheet Name**: Grouped by sheet for easy navigation
2. **Row/Column**: Sorted by position for logical flow
3. **Context**: Each difference shows related header and reference information

## Error Handling

- **File Not Found**: Re-prompts for valid file path
- **Invalid Format**: Only accepts Excel files (.xlsx, .xls, .xlsm)
- **Sheet Mismatch**: Reports sheets present in one file but not the other
- **Error Values**: Detects and reports Excel error formulas

## Project Structure

```
Excel comparator/
├── excel_comparator.py    # Main application
├── README.md              # This file
└── requirements.txt       # Python dependencies
```

## Author

**Shreyes Shalgar (shalsh1)**

## License

Internal use only
