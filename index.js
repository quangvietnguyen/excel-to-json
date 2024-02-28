const ExcelJS = require('exceljs');
const fs = require('fs');

function main(file) {
  const list = [];
  var workbook = new ExcelJS.Workbook();
  workbook.xlsx
    .readFile(file)
    .then(function () {
      workbook.eachSheet(function (worksheet, sheetId) {
        worksheet.eachRow(function (row, rowNumber) {
          rowNumber > 1 &&
            list.push({ appcode: row.values[1], appname: row.values[2] });
        });
      });
    })
    .finally(function () {
      var json = JSON.stringify(list);
      fs.writeFile('path/to/your/json/file.json', json, 'utf8', function (err) {
        if (err) throw err;
        console.log('complete');
      });
    });
}

main('path/to/your/excel/file.xlsx');
