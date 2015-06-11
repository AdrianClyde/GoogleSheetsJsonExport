// Add this code on the sheet you wish to export
// Tools > Script Editor > Copy paste the code in
// After saving go JSON Export > Export Sheets


// Menu code

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('JSON Export')
      .addItem('Export Sheets', 'menuItem1')
      .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi()
     start();
}


// Generating code below


var printed = "";

function start(){ 

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  if (sheets.length > 1) {
    for each(sheet in sheets) getInfo(sheet);
  }
  
  var finalPrint = printed.toString();
  // Remove final comma
  finalPrint = finalPrint.replace(/,([^,]*)$/, "<br>");
  
  var htmlOutput = HtmlService
  .createHtmlOutput('<p>{<br>' + finalPrint + '}</p>')
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setWidth(800)
  .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Paste into new .json file');
}

function getInfo(sheet) {
  // Find depth of columns and rows
  var depth = sheet.getLastRow() -1;
  var width = sheet.getLastColumn();
  
  // Get titles (key names)
  var k = sheet.getRange(1, 1, 1, width);
  var keys = k.getValues();
  //changing keys to 1 dimensional array
  keys = keys[0];
  
  var contents = [];
  
  for(var i=0; i < width; i++) {
    var range = sheet.getRange(2, i+1, depth, 1);
    var values = range.getValues();
    
    contents.push(values);
  }
  formatter(contents, sheet, keys);  
}

function formatter(values, sheet, keys) {
  var parentObjName = sheet.getSheetName();
  printed += '"' + parentObjName + '":[<br>';

  for(var x=0 ; x < values[0].length; x++){  
    printed += '&nbsp;&nbsp;&nbsp;&nbsp;{';
    for(var i=0 ; i < keys.length; i++){
      keys[i] = keys[i].replace(/"/g,'\\"');
      values[i][x][0] = values[i][x][0].replace(/"/g,'\\"');
      
      printed += '"' + keys[i] + '":"' + values[i][x][0] + '"';
      if(i<keys.length-1) printed += ', ';
    }    
    printed += '}';
    if(x < values[0].length -1) printed += ',<br>';
  }
  
  printed += '<br>],<br>';
}
