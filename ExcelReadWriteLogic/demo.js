const ExcelJS = require('exceljs');
//ITS SUPPORTS BOTH XLSX AND XLS AND ALSO CSV and JSON
// Create a new workbook




async function writeExcelFile(searchText, replaceText,chnage, filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath)

    const worksheet = workbook.getWorksheet('Sheet1');
  const output =   await readExcel(worksheet, searchText)

    const cell = worksheet.getCell(output.row, output.col+chnage.colChange)
    cell.value = replaceText
    await workbook.xlsx.writeFile(filePath)

    console.log(`File data is updated this -${searchText} to this ${replaceText} `)
}

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