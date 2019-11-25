const electron = require("electron");
var Excel;
var workbook;
var frequencies = {};
async function parse(filename, type, percentage, out) {
  Excel = require("exceljs");
  workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filename);
  var worksheet = workbook.getWorksheet(1);
  worksheet.eachRow(function(row, rowNumber) {
    if (rowNumber != 1) {
      row.eachCell(function(cell, colNumber) {
        title = worksheet.getRow(1).getCell(colNumber);
        if (frequencies[title] == undefined) {
          frequencies[title] = {};
        }
        if (frequencies[title][cell.text] == undefined) {
          frequencies[title][cell.text] = 1;
        } else {
          frequencies[title][cell.text]++;
        }
      });
    }
  });
  if (percentage != 0) {
    let percentageUnit = parseInt(percentage);
    for (const title of Object.keys(frequencies)) {
      let sum = 0;
      for (const field of Object.keys(frequencies[title])) {
        sum += frequencies[title][field];
      }
      let tempObj = {};
      tempObj[title] = {};
      for (const field of Object.keys(frequencies[title])) {
        tempObj[title][field] =
          (frequencies[title][field] / sum) * percentageUnit;
      }
      frequencies[title] = tempObj[title];
    }
  }
  makeCharts(type, out);
}
function makeCharts(type, out) {
  var fs = require("fs");
  var XLSXChart = require("xlsx-chart");
  var xlsxChart = new XLSXChart();
  var opts;
  var count = Object.keys(frequencies).length;
  for (const title of Object.keys(frequencies)) {
    if (title != "") {
      console.log(frequencies[title]);
      var data = {};
      data[title] = frequencies[title];
      opts = {
        chart: type,
        titles: [title],
        fields: Object.keys(frequencies[title]),
        data: data,
        chartTitle: title
      };
      xlsxChart.generate(opts, function(err, data) {
        if (err) {
          console.error(err);
        } else {
          fs.writeFileSync(
            out + "/" + title.substring(0, 50).replace("/", "") + ".xlsx",
            data
          );
          console.log(
            title.substring(0, 50).replace("/", "") + ".xlsx created."
          );
        }
        count--;
        if (count == 0) {
          process.exit(0);
        }
      });
    }
  }
}

const readline = require("readline").createInterface({
  input: process.stdin,
  output: process.stdout
});
readline.question("Insert Filename: ", filename => {
  readline.question(
    `Chart Type: (column, bar, line, area, radar, scatter, pie) `,
    type => {
      readline.question(`Percentage: (0 = None)`, percentage => {
        readline.question(`Output Folder: `, out => {
          parse(filename, type, percentage, out);
          readline.close();
        });
      });
    }
  );
});
