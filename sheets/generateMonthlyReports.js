/* eslint-disable no-undef */
/**
 * @OnlyCurrentDoc
 */

const colorPallete = ['#3ab795', '#a0e8af', '#9addac', '#93d1a8', '#86baa1', '#bad2b9', '#edead0', '#f6dd93', '#fbd675', '#ffcf56'];
const borderStyle = SpreadsheetApp.BorderStyle.SOLID;

function cleanup(spreadsheet) {
  const sheets = spreadsheet.getSheets();

  sheets.slice(1).forEach((sheet) => {
    spreadsheet.deleteSheet(sheet);
  });
}

function generateMonthlyReports(spreadsheet) {
  const maxRows = spreadsheet.getSheetByName('Sheet1').getLastRow();
  const range = spreadsheet.getRange(`A2:F${maxRows}`);
  const values = range.getValues();

  const names = [];

  const reports = {};

  const timeZone = spreadsheet.getSpreadsheetTimeZone();
  const headers = spreadsheet.getSheetByName('Sheet1').getRange('A1:F1').getValues();

  values.forEach((value) => {
    if (!names.includes(value[1])) {
      names.push(value[1]);
    }

    const valueDate = new Date(value[2]);
    const reportMonth = Utilities.formatDate(valueDate, timeZone, 'MM-yyyy');
    // eslint-disable-next-line no-unused-expressions
    !(reportMonth in reports) && (reports[reportMonth] = []);

    reports[reportMonth].push(value);
  });
  const reportMonths = Object.keys(reports);

  reportMonths.forEach((month) => {
    let reportMonthSheet = spreadsheet.getSheetByName(month);
    if (!reportMonthSheet) {
      spreadsheet.insertSheet(month);
      reportMonthSheet = spreadsheet.getSheetByName(month);
      headers[0].push('total hours');
      reportMonthSheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setFontWeight('bold');
    }
    const reportsData = reports[month];
    const rowCount = reportsData.length;
    const columnCount = reportsData[0].length;
    const reportMonthRange = reportMonthSheet.getRange(2, 1, rowCount, columnCount);
    reportMonthRange.setValues(reportsData);

    const column = reportMonthSheet.getRange('D2:E');

    column.setNumberFormat('hh:mm');

    const rows = reportMonthSheet.getLastRow();
    // eslint-disable-next-line no-plusplus
    for (let i = 2; i <= rows; i++) {
      reportMonthSheet.getRange(i, 7).setFormula(`=(E${i}-D${i})`);
    }

    const sumByNameTable = [];
    names.forEach((name) => {
      const sumByName = `=SUMIF(B2:B${rows}, "${name}", G2:G${rows})*24`;
      sumByNameTable.push([name, sumByName]);
    });
    const reportSumRangeHeader = reportMonthSheet.getRange(1, 10, 1, 2);
    reportSumRangeHeader.setValues([['Name', 'Total Hours per month']]).setFontWeight('bold');

    const reportSumRange = reportMonthSheet.getRange(2, 10, names.length, 2);
    reportSumRange.setValues(sumByNameTable);
    reportSumRange.setBorder(true, true, true, true, true, true, null, borderStyle);
    const backgrounds = reportSumRange.getBackgrounds();
    const newBckgs = backgrounds.map((row, index) => {
      const color = colorPallete[index % 9];
      // eslint-disable-next-line no-param-reassign
      row[0] = color;
      // eslint-disable-next-line no-param-reassign
      row[1] = color;
      return row;
    });
    reportSumRange.setBackgrounds(newBckgs);
  });
}

// eslint-disable-next-line no-unused-vars
function onOpen() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  cleanup(spreadsheet);
  generateMonthlyReports(spreadsheet);
}
