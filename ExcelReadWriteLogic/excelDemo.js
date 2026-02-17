const ExcelJS = require('exceljs');
//ITS SUPPORTS BOTH XLSX AND XLS AND ALSO CSV and JSON
// Create a new workbook

async function readExcelFile() {


const workbook = new ExcelJS.Workbook();
//this code is aysnchronous so we need to use async and await to read the file
//read a excel file
await workbook.xlsx.readFile('/Users/aayus/OneDrive/Desktop/Projects/ExcelTestSheet.xlsx')

    // Get a worksheet to the workbook
    const worksheet = workbook.getWorksheet('Sheet1');
    //get some data
    worksheet.eachRow((row, rowNumber) => {
        //pull  ;; out all the data in the row
        row.eachCell((cell, colNumber) => {
            console.log(cell.value);

        })
    })
}

readExcelFile();