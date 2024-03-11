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

async function writeExcelTest(searchText, replaceText, filePath) {
  const workbook = new excelJs.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet("Sheet1");
  const output = await readExcel(worksheet, searchText);

  const cell = worksheet.getCell(output.row, output.column);
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
  "Banana",
  "Stawberry",
  "C:/Users/Ljubisha/Downloads/excelDownloadTest.xlsx"
);
