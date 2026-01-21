# Excel Processor

## Overview
The Excel Processor is a Python program designed to read data from an Excel file, perform calculations, and write the results back to the same Excel file. It is built to be maintainable, extensible, and robust, with support for multiple sheets and data types.

## Features
- **Read Data**: Read data from specified sheets in an Excel file.
- **Perform Calculations**: Perform various calculations on the data.
- **Write Results**: Write the results back to the specified location in the Excel file.
- **Extensibility**: Easily add new calculations and functionality.
- **Error Handling**: Handle errors gracefully and provide meaningful error messages.
- **Multiple Sheets**: Support for multiple sheets and creating new sheets.
- **Data Types**: Handle multiple types of data.

## Installation
1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd finance-tools
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
1. **Import the `ExcelProcessor` class**:
   ```python
   from src.excel_processor import ExcelProcessor
   ```

2. **Initialize the `ExcelProcessor`**:
   ```python
   processor = ExcelProcessor("path/to/your/file.xlsx")
   ```

3. **Define the calculations**:
   ```python
   calculations = {
       "Value1": "sum",
       "Value2": "average"
   }
   ```

4. **Process the Excel file**:
   ```python
   processor.process("InputSheet", "OutputSheet", calculations)
   ```

## Architecture
The program is divided into several modules:
- **ExcelProcessor**: Main class to orchestrate the process.
- **ExcelReader**: Handles reading data from Excel files.
- **Calculator**: Performs calculations on the data.
- **ExcelWriter**: Writes results back to the Excel file.
- **ErrorHandler**: Manages errors and logging.

## Testing
To run the tests, use the following command:
```bash
python -m unittest tests.test_excel_processor
```

## Documentation
- Calculation Elements v0
  - Iteration unit
  - Qualifiers (e.g. must-have, must-not-have, lookup)
  - Computation (using numbers within the unit + elsewhere)
  - Accumulation (add or subtract)
  - Final result computation if any
  - Validation or sanity checking
  - Handling missing data (blank numbers; missing required elements such as qualifiers; missing units - making right assumptions and getting them validated)
  - Handling comments / annotations (e.g. uncertainty, conditions, etc.)
  - Logging

## License
This project is licensed under the MIT License.