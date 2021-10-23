// constants
const SHEET = SpreadsheetApp.getActiveSheet();
const LOG_IT = true;
const HEADER_ROW = 1;

function logIt(){
  if( !LOG_IT ){
    return
  }
  for (let i = 0; i < arguments.length; i++) {
    console.log(arguments[i]);
  }
}

/**
 * Write Array to current Sheets
 * @param arr {array} - the 2-dim array with data
 * @param startRow {integer} - row to start at (optional)
 * @param startCol {integer} - column to start at (optional)
 */
function writeArrayToActiveSheet(arr, startRow=1, startCol=1){
  if( !arr.isArray() ) return
  return SHEET.getRange( startRow, startCol, arr.length, arr[0].length ).setValues( arr )
}

/**
 * Write array for one row to Sheet
 * @param {Object} Sheet
 * @param {Array} Array - 2-Dim Array [[cell1, cell2, …]]
 * @param {Number} RowNum
 * @param {Number} ColNum
 */
function writeArrayToSheet(Sheet, arr, startRow=1, startCol=1){
  //if( !arr.isArray() ) return
  return Sheet.getRange(startRow, startCol, 1, arr[0].length).setValues(arr);
}

/**
* Checks to see if we are in the correct sheet
* @param {Object} Active Sheet - The curently active Sheet 
* @param {String} Sheet Name - The name of the sheet to test against
* @returns {Boolean} Is Active = Target
*/
function isSheetCorrect( currentSheet, targetSheetName ){
  return currentSheet.getSheetName() == targetSheetName
}


/**
* Extract G-Drive Id from Url
* 
* @param {string} A url to the G-Drive File
* @returns {string} The id
*/
function getFileIdFromFileUrl( url ) { 
  return url.match( /[-\w]{25,}/ )
}


/**
* Returns header as objects with colName : colNum
* 
* @param {Object} Sheet which is used
* @param {number} The header row
* @returns {Object} ColName: ColNum
*/
function getHeaderAsObjFromSheet( sheet, headerRow = 1 ){
  const headerArr = sheet.getRange( headerRow, 1, 1, sheet.getLastColumn() ).getValues()[0];  
  return convertArrToObj_( headerArr );
}

/**
* Returns an array of objects for each selected row headerName : rowValue
* Assumes a continues range is selected, ie no hidden, filtered or multiple ranges
* 
* @param {Object} Sheet - The currently active Sheet
* @param {number} Header Row - The row where the header is located
* @return {Array} Collection of rows as objects with rowNum
*/
function getSelectedRowsAsObjInArr( sheet, headerRow ){  
  const firstRow = sheet.getActiveRange().getRow();
  const lastRow  = sheet.getActiveRange().getLastRow();  
  const lastCol  = sheet.getLastColumn();
  const rowsAsArr = sheet.getRange( firstRow, 1, lastRow - firstRow + 1, lastCol ).getValues();
  const headerObj = getHeaderAsObjFromSheet(sheet, headerRow);
  
  const rowsAsObjInArr = [];
  rowsAsArr.forEach( (row, i) => rowsAsObjInArr.push( HELPER.convertRowArrToObj_(row, headerObj) ));
  rowsAsObjInArr.forEach( (rowObj, i) => rowObj["rowNum"] = firstRow + i );
  return rowsAsObjInArr
}

/**
* Returns an array of objects for the entire sheet's rows headerName : rowValue
* Assumes a continues range is selected, ie no hidden, filtered or multiple ranges
* 
* @param {Object} Sheet - The currently active Sheet
* @param {number} Header Row - The row where the header is located
* @return {Array} Collection of rows as objects with rowNum
*/
function getAllRowsAsObjInArr( sheet, headerRow ){  
  const rowsAsArr = sheet.getDataRange().getValues();
  const headerObj = getHeaderAsObjFromSheet(sheet, headerRow);
  
  const rowsAsObjInArr = [];
  rowsAsArr.forEach( (row, i) => rowsAsObjInArr.push( HELPER.convertRowArrToObj_(row, headerObj) ));
  rowsAsObjInArr.forEach( (rowObj, i) => rowObj["rowNum"] = ++i );
  return rowsAsObjInArr
}

/**
* Returns an Object where { headerName : rowValue}
* Assumes a continues range is selected, ie no hidden, filtered or multiple ranges
* 
* @param {Object} Sheet - The currently active Sheet
* @param {Number} Header Row - The row where the header is located
* @param {Number} RowNum - Number of row to get
* @return {Object} Row as Obj + rowNum
*/
function getRowAsObj( sheet, headerRow, rowNum ){
  const rowsAsArr = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerObj = getHeaderAsObjFromSheet(sheet, headerRow);

  return HELPER.convertRowArrToObj_(rowsAsArr, headerObj);
}



/**
* Appends the array to the target Sheet
* @param {array} Data Array - 2-d array of data to write
* @param {string} Target Sheet Name - Where the data should be appended to
*/
function appendRowsToTargetSheet( dataArray, sheetName ){
    const outputSheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    const outputRange = outputSheet.getRange(outputSheet.getLastRow()+1, 1, dataArray.length, dataArray[0].length);
    outputRange.setValues(dataArray)
}

function headerColFun( header, range ){
  return range.indexOf(header) + 1     
}

/**
 * Writes to the current Sheet
 * @param {Integer} RowNum - Row, 1-index
 * @param {Integer} ColNum - Column, 1-index
 * @param {Integer} Value - Inserted into the cell
 */
function writeToSheet( row, column, value ){
  SHEET.getRange( row, column ).setValue( value )
}

/**
 * Writes to the current Sheet
 * @param {Object} Sheet - The Sheet to write to
 * @param {Integer} RowNum - Row, 1-index
 * @param {Integer} ColNum - Column, 1-index
 * @param {Integer} Value - Inserted into the cell
 * @return {Object} Sheets
 */
function writeToOtherSheet( sheet, row, column, value ){
  return sheet.getRange( row, column ).setValue( value )
}

/**
 * Returns the first empty row in a particular column
 * @param {Object} Sheet - The sheet to use
 * @param {Integer} ColNum - The column to use
 * @return {Integer} RowNum - First free row in said column
 */
function getFirstEmptyRowInCol( sheet, colNum){
  const allRowsInSheet = sheet.getRange(1, colNum, sheet.getLastRow() ).getValues();
  const lastRowInThisCol = allRowsInSheet.filter(String).length;
  return lastRowInThisCol + 1
}

function getHeader_(sheet, headerRow){
  var headerArr = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];  
  var headerObj= convertArrToObj_(headerArr);
  return headerObj
}

function getSelectedRowsAsObjInArr_( headerObj ){
  var firstRow = SHEET.getActiveRange().getRow();
  var lastRow  = SHEET.getActiveRange().getLastRow();  
  var rowsAsArr = SHEET.getRange(firstRow, 1, lastRow - firstRow + 1, SHEET.getLastColumn()).getValues();
  
  var rowsAsObjInArr = [];
  rowsAsArr.forEach(function( rowArr, index ){
    rowsAsObjInArr.push( convertRowArrToObj_(rowArr, headerObj) );
  });
  return rowsAsObjInArr
}

/**
+ Simple message alert to user
* @param {string} Message to display to user
*/
function alert( msg ){
  SpreadsheetApp.getUi().alert(msg)
}

const HELPER = (function(){
  
  function convertArrToObj_( anArray ){
    const asObj = {};
    anArray.forEach( (item, index) => {
                    asObj[item] = index + 1
                    });
    return asObj
  }
  
  function convertRowArrToObj_( rowAsArray, headerObj ){
    var rowAsObj = {};
    for(var header in headerObj){
      rowAsObj[header] = rowAsArray[ headerObj[header]-1 ]
    }
    return rowAsObj
  }
  
  function todayISO_(){
    return new Date().toISOString().substr(0,10)
  }
  
  function convertDateToISO_( aDate ){
    return aDate.toISOString().substr(0,10)
  }
  
  return {
    convertArrToObj_: convertArrToObj_,
    convertRowArrToObj_: convertRowArrToObj_,
    todayISO_: todayISO_,
    convertDateToISO_: convertDateToISO_
  }
  
})();
// moved to HELPER
function convertArrToObj_( anArray ){
  var asObj = {};
  anArray.forEach(function(item, index){
    asObj[item] = index + 1
  });
  return asObj
}

function convertRowArrToObj_( rowAsArray, headerObj ){
  var rowAsObj = {};
  for(var header in headerObj){
    rowAsObj[header] = rowAsArray[ headerObj[header]-1 ]
  }
  return rowAsObj
}

function todayISO(){
  return new Date().toISOString().substr(0,10)
}

Date.prototype.addHours = function(h) {
  this.setTime(this.getTime() + (h*60*60*1000));
  return this;
}
