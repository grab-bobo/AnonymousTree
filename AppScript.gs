/*  */
/* App Script functions and notes */

/* Connect to spreadsheet */
/*
  if sheetName is not null, get sheet by sheet name
  else, get active sheet 
*/ 
function connectSpreadsheet_bobo(sheetName){
  var app = SpreadsheetApp; // Connect to Spreadsheet
  var curSpreadsheet = app.getActiveSpreadsheet(); // Get current spreadsheet
  if (!Boolean(sheetName)) // ! null = true
    var activeSheet = curSpreadsheet.getSheetByName(sheetName); // Get sheet by name
  else // ! null = false
    var activeSheet = curSpreadsheet.getActiveSheet(); // Get current active sheet

  return activeSheet; // Return the connection
}

/*
  Get sheet's last row number
*/
function getLastRow_bobo(connection){
  var sheetLastRow = connection.getLastRow();
  return sheetLastRow;
}

/*
  Get sheet's last column number
*/
function getLastColumn_bobo(connection){
  var sheetLastColumn = connection.getLastColumn();
  return sheetLastColumn;
}

/*
  Get whole data table in sheet
*/
function getDataRange_Bobo(connection){
  var data = connection.getDataRange().getValues();
  return data;
}

/*
  Get specific cell value in sheet
*/
/*
  if isNotation is true, range by notation
  else, range by number coordinates, w = x-axis, x = y-axis, y = no of rows, z = no of columns
*/ 
function getCellValue_bobo(connection, isNotation ,w, x, y, z){
  if (Boolean(isNotation)) // isNotation = true
    var cellValue = connection.getRange(w + ":"+ x).getValues(); // range by notation
  else // isNotation = false
    var cellValue = connection.getRange(w, x, y, z).getValues(); // range by number coordinates
  return cellValue;

}

function main_bobo(){
  var connection = connectSpreadsheet_bobo("sheet1");
  var sheetLastRow = getLastRow_bobo(connection);
  var sheetLastColumn = getLastColumn_bobo(connection);
  var values = getCellValue_bobo(connection, false, 1, 1, sheetLastRow, sheetLastColumn);
  var data = getDataRange_Bobo(connection);
  Logger.log(values)
  Logger.log(sheetLastRow);
  Logger.log(data);
}
