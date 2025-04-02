## Spreadsheet Automation

This repository provides a Python-based automation tool to process spreadsheet files, specifically Excel workbooks. It is designed to perform tasks like adjusting prices and generating charts on the first sheet of the workbook. The automation is built using Python and leverages the `openpyxl` library to interact with Excel files.

## Features

- **Adjusting values**: The `process_workbook` function multiplies values in the third column (starting from the second row) by 0.9 and updates the fourth column with the corrected values.
- **Chart Generation**: The module automatically creates a bar chart based on the corrected values in column 4 and places it on the sheet.
- **Error Handling**: The code handles common errors, including file not found and missing sheets, with appropriate messages.

## Requirements

- Python 3.x
- `openpyxl` library

You can install the required dependencies by running the following command:

```bash
pip install openpyxl
```

## Directory Structure

```
├── ProcessSpreadSheet/
│   ├── __init__.py
│   └── processor.py    # Contains the process_workbook function
├── app.py              # Main application script
└── README.md           # Project documentation
```

## How to Use

1. Clone this repository to your local machine:

   ```bash
   git clone https://github.com/yourusername/spreadsheet-automation.git
   cd spreadsheet-automation
   ```

2. Make sure the Excel file you want to process is in the same directory and provide its full name with dot extension.

3. Run the `app.py` file:

   ```bash
   python app.py
   ```

4. You will be prompted to enter the filename of the Excel workbook you want to process. After entering the filename, the script will process the first sheet (`Sheet1`), adjust the prices in the third column, create a bar chart, and save the modified workbook.

## Code Overview

### `ProcessSpreadSheet/processor.py`

This file contains the `process_workbook` function, which does the following:

- Loads the workbook.
- Processes the first sheet (`Sheet1`).
- Adjusts the values in the third column by multiplying them by 0.9.
- Generates a bar chart based on the corrected values in column 4.
- Saves the modified workbook.

#### Example `process_workbook` Function:

```python
from sys import exception
import openpyxl as xl
from openpyxl.chart import BarChart, Reference

def process_workbook(filename):
    try:
        wb = xl.load_workbook(filename)
        sheet = wb['Sheet1']
        for row in range(2, sheet.max_row + 1):
            cell = sheet.cell(row, 3)
            corrected_value = cell.value * 0.9
            corrected_price_cell = sheet.cell(row, 4)
            corrected_price_cell.value = corrected_value

        values = Reference(sheet, min_row=2, max_row=sheet.max_row, min_col=4, max_col=4)
        chart = BarChart()
        chart.add_data(values)

        sheet.add_chart(chart, 'e2')
        wb.save(filename)
    except FileNotFoundError:
        print("Entered File Name does not exist in directory")
    except KeyError:
        print("Worksheet does not exist.")
```

### `app.py`

This file contains the main application script. It prompts the user for the Excel file name and invokes the `process_workbook` function from the `ProcessSpreadSheet` module.

```python
from ProcessSpreadSheet import process_workbook

inputFileName = input("Enter File Name: ")
process_workbook(inputFileName)
```

## Error Handling

- **File Not Found**: If the specified file does not exist in the directory, an error message is displayed.
- **Missing Worksheet**: If the worksheet `Sheet1` is not found, an error message is displayed.

