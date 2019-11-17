var file = "./in.xlsx";
var Excel;
var workbook;
var frequencies = {};
async function parse() {
  Excel = require("exceljs");
  workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(file);
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
  makeCharts();
}
function makeCharts() {
  var fs = require("fs");
  var XLSXChart = require("xlsx-chart");
  var xlsxChart = new XLSXChart();
  var opts;
  for (const title of Object.keys(frequencies)) {
    console.log(frequencies[title]);
    console.log(Object.keys(frequencies[title]));
    var data = {};
    data[title] = frequencies[title];
    opts = {
      chart: "pie",
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
          title.substring(0, 50).replace("/", "") + ".xlsx",
          data
        );
        console.log(title.substring(0, 50).replace("/", "") + ".xlsx created.");
      }
    });
  }
}
parse();
