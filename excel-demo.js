const excelJs = require("exceljs");

// const workbook = new excelJs.Workbook();
// workbook.xlsx.readFile('C:/Users/Ljubisha/Downloads/excelDownloadTest.xlsx').then(function()
// {
//     const worksheet = workbook.getWorksheet('Sheet1');
//     worksheet.eachRow((row,rowNumber) =>
//     {
//         row.eachCell((cell, colNumber) =>
//         {
//             console.log(cell.value);
//         })
//     })

// })

async function writeExcelTest(searchText, replaceText, change, filePath) {
  const workbook = new excelJs.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet("Sheet1");
  const output = await readExcel(worksheet, searchText, change);

  const cell = worksheet.getCell(output.row, output.column+change.colChange);
  cell.value = replaceText;
  await workbook.xlsx.writeFile(filePath);
}

async function readExcel(worksheet, searchText) {
  let output = { row: -1, column: -1 };
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      if (cell.value === searchText) {
        output.row = rowNumber;
        output.column = colNumber;
      }
    });
  });
  return output;
}

writeExcelTest(
  "Apple",
  999,
  {rowChange:0, colChange:2},
  "C:/Users/Ljubisha/Downloads/excelDownloadTest.xlsx"
);
