//https://www.benlcollins.com/apps-script/how-to-connect-google-sheets-to-a-database/
/**
 * @OnlyCurrentDoc
 */

var MAXROWS = 1000
var SEEKWELL_J_SHORT_DATES = { day: "yyyy-MM-dd", month: "yyyy-MM", year: "yyyy", dayNum: "dd", monthNum: "MM", yearNum: "yyyy", week: "W" }
var SEEKWELL_J_TIMEZONE = "UTC"
var HOST = 'INSERT DB ADDRESS'
var PORT = '3306'
var USERNAME = 'INSERT USERNAME'
var PASSWORD = 'INSERT PASSWORD'
var DATABASE = 'INSERT DB NAME'
var DB_TYPE = 'mysql'

function goToSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName(sheetName));
};

function runSql(query, options) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getActiveSheet();
  var sheetName = sheet.getName();
  var cell = doc.getActiveSheet().getActiveCell();
  var activeCellRow = cell.getRow();
  var activeCellCol = cell.getColumn();
  try {
    var fullConnectionString = 'jdbc:' + DB_TYPE + '://' + HOST + ':' + PORT
    var conn = Jdbc.getConnection(fullConnectionString, USERNAME, PASSWORD);
    console.log('query :', query)
    var stmt = conn.createStatement();
    stmt.execute('USE ' + DATABASE);
    var start = new Date();
    var stmt = conn.createStatement();
    stmt.setMaxRows(MAXROWS);
    var rs = stmt.executeQuery(query);
  } catch (e) {
    console.log(e, e.lineNumber);
    Browser.msgBox(e);
    return false;
  }
  var results = [];
  cols = rs.getMetaData();
  console.log("cols", cols)
  var colNames = [];
  var colTypes = {};
  for (i = 1; i <= cols.getColumnCount(); i++) {
    var colName = cols.getColumnLabel(i)
    colTypes[colName] = { type: cols.getColumnTypeName(i), loc: i }
    colNames.push(colName);
  }
  var rowCount = 1;
  results.push(colNames);
  while (rs.next()) {
    curRow = rs.getMetaData();
    rowData = [];
    for (i = 1; i <= curRow.getColumnCount(); i++) {
      rowData.push(rs.getString(i));
    }
    results.push(rowData);
    rowCount++;
  }
  rs.close();
  stmt.close();
  conn.close();
  console.log('results', results)
  var colCount = results[0].length
  var rowCount = results.length
  var comment = "Updated on: " + (new Date()) + "\n" + "Query:\n" + query
  if (options.omitColumnNames) {
    results = results.slice(1)
    rowCount -= 1
  }
  if (options.clearColumns && sheet.getLastRow() > 0) {
    var startCellRange = sheet.getRange(startCell)
    sheet.getRange(startCellRange.getRow(), startCellRange.getColumn(), sheet.getLastRow(), colCount).clearContent();
  }
  if (options.clearSheet) {
    var startCellRange = sheet.getRange(startCell)
    sheet.clear({ contentsOnly: true });
  }
  sheet.getRange(activeCellRow, activeCellCol, rowCount, colCount).clearContent();
  sheet.getRange(activeCellRow, activeCellCol, rowCount, colCount).setValues(results);
  var cell = sheet.getRange(activeCellRow, activeCellCol)
  cell.clearNote()
  cell.setNote(comment);
  sheet.setActiveRange(sheet.getRange(activeCellRow + rowCount + 1, activeCellCol))
  console.log('query success!, rows = ', rowCount - 1)
}

function arrayToCommaSeparatedString(array) {
  var arrayString = "";
  if (array.length > 0)
    arrayString += array[0];
  else
    return "";

  for (var i = 1; i < array.length; i++)
    arrayString += ", " + array[i];

  return arrayString;
}

function arrayToCommaSeparatedQuotedString(array) {
  var arrayString = "";
  if (array.length > 0)
    arrayString += "'" + array[0] + "'";
  else
    return "";

  for (var i = 1; i < array.length; i++)
    arrayString += ", " + "'" + array[i] + "'";

  return arrayString;
}

function runInsert(query, options) {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getActiveSheet();
  var sheetName = sheet.getName();
  var cell = doc.getActiveSheet().getActiveCell();
  var activeCellRow = cell.getRow();
  var activeCellCol = cell.getColumn();
  var value = cell.getDisplayValue();

  var selection = sheet.getSelection();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRange(selection.getActiveRange().getA1Notation());
  //console.log(range)
  var values = range.getDisplayValues();
  var cols = ['answ', 'answ2'];
  var colsString = arrayToCommaSeparatedString(cols);

  try {
    var fullConnectionString = 'jdbc:' + DB_TYPE + '://' + HOST + ':' + PORT
    var conn = Jdbc.getConnection(fullConnectionString, USERNAME, PASSWORD);
    console.log('query :', query)
    var stmt = conn.createStatement();
    stmt.execute('USE ' + DATABASE);
    var start = new Date();
    var stmt = conn.createStatement();
    stmt.setMaxRows(MAXROWS);

    for (var i = 0; i < values.length; i++) {  //iterate over rows
      var rowString = arrayToCommaSeparatedQuotedString(values[i]);
      stmt.executeUpdate('INSERT INTO innodb.table (' + colsString + ') VALUES' + ' (' + rowString + ');');
    }
  } catch (e) {
    console.log(e, e.lineNumber);
    Browser.msgBox(e);
    return false;
  }
  stmt.close();
  conn.close();
}

function runA2QueryFromSheet() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sql = doc.getRange('query!a2').getDisplayValue();
  var options = {}
  Logger.log('sql;', sql)
  runSql(sql, options)
}

function runA3QueryFromSheet() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sql = doc.getRange('query!a3').getDisplayValue();
  var options = {}
  Logger.log('sql;', sql)
  runInsert(sql, options)
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Database')
    .addItem('Get data', 'runA2QueryFromSheet')
    .addItem('Push data', 'runA3QueryFromSheet')
    .addToUi();
}

function launch() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Database');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}
