const ExcelJS = require('exceljs');
// High-level flow: load workbook -> find a value -> offset the cell -> update -> save




// Update a cell located by searchText, then shifted by change offsets
async function writeExcelFile(searchText, replaceText,chnage, filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath)

    // Locate the sheet and find the coordinates of searchText
    const worksheet = workbook.getWorksheet('Sheet1');
    const output =   await readExcel(worksheet, searchText)

    // Apply column offset before writing the new value
    const cell = worksheet.getCell(output.row, output.col+chnage.colChange)
    cell.value = replaceText
    await workbook.xlsx.writeFile(filePath)

    console.log(`File data is updated this -${searchText} to this ${replaceText} `)
}

// Scan the sheet to find the first cell that matches searchText
async function readExcel(worksheet, searchText) {
let output = { row: -1, col: -1 };
    worksheet.eachRow((row, rowNumber) => {

        row.eachCell((cell, colNumber) => {
            if (cell.value === searchText) {

                output.row = rowNumber;
                output.col = colNumber;
            };

        })
    })

    return output;
}



writeExcelFile("80000", "9",{rowchage:0, colChange:2}, "/Users/aayus/OneDrive/Desktop/Projects/ExcelTestSheet.xlsx");