# ExcelUtils

Excel read and write operations using ExcelJS: includes examples for reading data, finding cells by value, updating cells with column offsets, and dynamic search-and-replace functionality.

## Overview

This project demonstrates how to work with Excel files using the ExcelJS library in Node.js. It includes multiple examples for common operations:
- Reading data from Excel sheets
- Finding cells by value
- Updating cells with row/column offsets
- Search and replace with dynamic offsets

## Dependencies

- `exceljs`: ^4.4.0 - Excel file processing library supporting XLSX, XLS, CSV, and JSON formats

## Setup

Install dependencies:
```bash
npm install
```

## Running Examples

Update the file paths in each script to match your Excel file location, then run:

```bash
# Read Excel file
node ExcelReadWriteLogic/excelDemo.js

# Write/Search with offset
node ExcelReadWriteLogic/ReadWriteDynamicSearch.js

# Search and replace (hardcoded value)
node ExcelReadWriteLogic/WriteDataExcelSenario2.js

# Search and replace (with offset)
node ExcelReadWriteLogic/demo.js
```

## Line-by-Line Explanations

### ReadWriteDynamicSearch.js

The main flow: load workbook -> find a value -> offset the cell -> update -> save.

1. `const ExcelJS = require('exceljs');` - Import ExcelJS library
2. `async function writeExcelFile(searchText, replaceText, chnage, filePath)` - Function to find and update a cell with optional offset
3. `const workbook = new ExcelJS.Workbook();` - Create a new workbook instance
4. `await workbook.xlsx.readFile(filePath)` - Load Excel file asynchronously
5. `const worksheet = workbook.getWorksheet('Sheet1');` - Get the worksheet named 'Sheet1'
6. `const output = await readExcel(worksheet, searchText)` - Call helper function to locate the search text
7. `const cell = worksheet.getCell(output.row, output.col+chnage.colChange)` - Get cell at found coordinates, plus column offset
8. `cell.value = replaceText` - Set new value to the cell
9. `await workbook.xlsx.writeFile(filePath)` - Save workbook back to file
10. `async function readExcel(worksheet, searchText)` - Helper function to scan and find a value
11. `worksheet.eachRow((row, rowNumber) => { ... })` - Iterate through each row
12. `row.eachCell((cell, colNumber) => { ... })` - Iterate through each cell in row
13. `if (cell.value === searchText)` - Check if cell matches search text
14. `output.row = rowNumber; output.col = colNumber;` - Store the coordinates when found
15. `return output;` - Return the found coordinates

### excelDemo.js

Basic read example that prints all cell values.

1. `const ExcelJS = require('exceljs');` - Import ExcelJS
2. `async function readExcelFile()` - Async function to read sheet
3. `const workbook = new ExcelJS.Workbook();` - Create workbook
4. `await workbook.xlsx.readFile(filePath)` - Read file into memory (I/O is async)
5. `const worksheet = workbook.getWorksheet('Sheet1');` - Get worksheet
6. `worksheet.eachRow((row, rowNumber) => { ... })` - Iterate rows
7. `row.eachCell((cell, colNumber) => { ... })` - Iterate cells
8. `console.log(cell.value);` - Print each cell value
9. `readExcelFile();` - Execute the function

### WriteDataExcelSenario2.js

Find a specific hardcoded value and update it with a new value.

1. `let output = {row:-1, col:-1};` - Initialize output object to track found cell
2. `async function writeExcelFile()` - Function to search and replace
3. `const workbook = new ExcelJS.Workbook();` - Create workbook
4. `await workbook.xlsx.readFile(filePath)` - Load file
5. `const worksheet = workbook.getWorksheet('Sheet1');` - Get worksheet
6. `worksheet.eachRow(...)` - Loop through rows
7. `if (cell.value === "Apple")` - Check for hardcoded value "Apple"
8. `output.row = rowNumber; output.col = colNumber;` - Store coordinates
9. `const cell = worksheet.getCell(output.row, output.col)` - Get the found cell
10. `cell.value = ("Aayush Mishra")` - Update with new value
11. `await workbook.xlsx.writeFile(filePath)` - Save file
12. `writeExcelFile();` - Execute

### demo.js

Similar to ReadWriteDynamicSearch.js but without comments. Demonstrates search with offset capability.

## Ideas and Improvements

1. **Add error handling**: Check if the searched value is found before updating (avoid updating -1,-1 coordinates)
2. **Support multiple sheets**: Allow specifying sheet name as parameter instead of hardcoding 'Sheet1'
3. **Batch operations**: Add ability to find and replace multiple values in one operation
4. **CSV/JSON support**: ExcelJS also supports CSV and JSON, could add examples
5. **Fix typo in variable names**: Rename `chnage` to `change` and `rowchage` to `rowChange` for clarity
6. **Create a utility module**: Extract common functions into a reusable module
7. **Add logging**: Implement better logging for debugging and monitoring
8. **Support formulas**: Add examples of reading/writing formulas, not just values
9. **Styling**: Add examples of formatting cells (color, font, borders)
10. **Validation**: Add input validation for file paths and sheet names

## Notes

- All scripts hardcode file paths. Update these to match your system's paths before running.
- The library works with async/await for file I/O operations - all read/write operations are asynchronous.
- Make sure the Excel file has a sheet named 'Sheet1' or update the sheet name in the code.
- ExcelJS supports XLSX, XLS, CSV, and JSON formats automatically based on file extension.
