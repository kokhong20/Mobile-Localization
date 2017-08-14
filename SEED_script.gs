// Includes functions for exporting active sheet or selected data as XML or Strings object
// to xml/strings file in the Google Drive folder where current spreadsheet located.

var HEADER_LIST = [];
/* Defaults for this particular spreadsheet, change as desired */
var DEFAULT_STRING;
var DEFAULT_VALUE;
var DEFAULT_TYPE = "Android";
var EXPORT_TYPE = ["Android","iOS"];

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    {name: "Configure export", functionName: "exportOptions"},
  ];
  ss.addMenu("Export XML/Strings", menuEntries);
}
    
    
function exportOptions() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Export XML/Strings');
  
  var sheet = doc.getActiveSheet();
  var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];
  HEADER_LIST = normalizeHeaders_(headers);
  DEFAULT_STRING = HEADER_LIST[0];
  if (HEADER_LIST.length > 1)
    DEFAULT_VALUE = HEADER_LIST[1];
  else
    DEFAULT_VALUE = HEADER_LIST[0];
  
  var grid = app.createGrid(8, 1);
  grid.setWidget(0, 0, makeLabel(app, 'String Name:'));
  grid.setWidget(1, 0, makeListBox(app, 'string', HEADER_LIST));
  grid.setWidget(2, 0, makeLabel(app, 'String Value:'));
  grid.setWidget(3, 0, makeListBox(app, 'value', HEADER_LIST));
  grid.setWidget(4, 0, makeButton(app, grid, 'Export Selected Data', 'exportSelectedRange'));
  grid.setWidget(5, 0, makeButton(app, grid, 'Export Active Sheet', 'exportSheet'));
  grid.setWidget(6, 0, makeLabel(app, 'Export type:'));
  grid.setWidget(7, 0, makeListBox(app, 'type', EXPORT_TYPE));
  app.add(grid);
  
  doc.show(app);
}

function makeLabel(app, text, id) {
  var lb = app.createLabel(text);
  if (id) lb.setId(id);
  return lb;
}

function makeListBox(app, name, items) {
  var listBox = app.createListBox().setId(name).setName(name);
  listBox.setVisibleItemCount(1);
  
  var cache = CacheService.getPublicCache();
  var selectedValue = cache.get(name);
  Logger.log(selectedValue);
  for (var i = 0; i < items.length; i++) {
    listBox.addItem(items[i]);
    if (items[1] == selectedValue) {
      listBox.setSelectedIndex(i);
    }
  }
  return listBox;
}

function makeButton(app, parent, name, callback) {
  var button = app.createButton(name);
  app.add(button);
  var handler = app.createServerClickHandler(callback).addCallbackElement(parent);;
  button.addClickHandler(handler);
  return button;
}

function makeTextBox(app, name) { 
  var textArea    = app.createTextArea().setWidth('100%').setHeight('200px').setId(name).setName(name);
  return textArea;
}

function exportSheet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowsData = getRowsData_(sheet, getExportOptions(e));
  var json = makeJSON_(rowsData);
  var options  = getExportOptions(e);
  if(options.type == "Android"){
    return saveAsXml(sheet, json, getExportOptions(e));
  }else{
    return saveAsStrings(sheet, json, getExportOptions(e));
  }
}

function exportSelectedRange(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowsData = getSelectedRowsData_(sheet);
  var json = makeJSON_(rowsData);
  var options  = getExportOptions(e);
  if(options.type == "Android"){
    return saveAsXml(sheet, json, getExportOptions(e));
  }else{
    return saveAsStrings(sheet, json, getExportOptions(e));
  }
}

function getExportOptions(e) {
  var options = {};
  
  options.string = e && e.parameter.string || DEFAULT_STRING;
  options.value = e && e.parameter.value || DEFAULT_VALUE;
  options.type = e && e.parameter.type || DEFAULT_TYPE;
  
  var cache = CacheService.getPublicCache();
  cache.put('string',   options.string);
  cache.put('value',   options.value);
  cache.put('type', options.type);
  
  Logger.log(options);
  return options;
}

function makeJSON_(object) {
  var jsonString = JSON.stringify(object, null, 4);
  return jsonString;
}

function saveAsXml(sheet, content, options) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // get the folder this spreadsheet locates in
  var thisFileId = SpreadsheetApp.getActive().getId();
  var thisFile = DriveApp.getFileById(thisFileId);
  var folder = thisFile.getParents().next();
  // append ".xml" extension to the sheet name
  var d = new Date();
  fileName = "Android_" + sheet.getName() + '_' + dateFormat(d) + ".xml";
  var jsonObject = Utilities.jsonParse(content);
  var xmlContent = json2xml(jsonObject, options);
  Logger.log("xmlContent:"+xmlContent);
  // create a file in the Folder with the given name and the xml data
  folder.createFile(fileName, xmlContent);
}

function saveAsStrings(sheet, content, options) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // get the folder this spreadsheet locates in
  var thisFileId = SpreadsheetApp.getActive().getId();
  var thisFile = DriveApp.getFileById(thisFileId);
  var folder = thisFile.getParents().next();
  // append ".strings" extension to the sheet name
  var d = new Date();
  fileName = "iOS_" + sheet.getName() + '_' + dateFormat(d) + ".strings";
  var jsonObject = Utilities.jsonParse(content);
  var stringsContent = json2strings(jsonObject, options);
  Logger.log("stringsContent:"+stringsContent);
  // create a file in the Folder with the given name and the strings data
  folder.createFile(fileName, stringsContent);
}

function dateFormat(d) {
  var fileName = d.getFullYear();
  if(d.getMonth()+1 < 10){
    fileName += '0' + (d.getMonth()+1)
  }else{
    fileName += '' + (d.getMonth()+1)
  }
  if(d.getDate() < 10){
    fileName += '0' + d.getDate()
  }else{
    fileName += '' + d.getDate()
  }
  if(d.getHours() < 10){
    fileName += '0' + d.getHours()
  }else{
    fileName += '' + d.getHours()
  }
  if(d.getMinutes() < 10){
    fileName += '0' + d.getMinutes()
  }else{
    fileName += '' + d.getMinutes()
  }
  return fileName;
}
//get active ranges as the data
function getSelectedRowsData_(sheet) {
  var headersRange = sheet.getRange(1, sheet.getActiveRange().getColumn() , sheet.getFrozenRows(), sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];
  var dataRange = sheet.getActiveRange();
  var objects = getObjects_(dataRange.getValues(), normalizeHeaders_(headers));
  return objects;
}
// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData_(sheet) {
  var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];
  var dataRange = sheet.getRange(sheet.getFrozenRows()+1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  var objects = getObjects_(dataRange.getValues(), normalizeHeaders_(headers));
  return objects;
}

// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - rowHeadersColumnIndex: specifies the column number where the row names are stored.
//       This argument is optional and it defaults to the column immediately left of the range; 
// Returns an Array of objects.
function getColumnsData_(sheet, range, rowHeadersColumnIndex) {
  rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
  var headersTmp = sheet.getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1).getValues();
  var headers = normalizeHeaders_(arrayTranspose_(headersTmp)[0]);
  return getObjects(arrayTranspose_(range.getValues()), headers);
}


// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects_(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders_(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader_(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader_(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum_(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit_(char) {
  return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose_(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}

function json2xml(o, options) {
   Logger.log("json");
  Logger.log(o);
  var xml = "<resource>";
  for(var m in o){
    var subObj = o[m];
    xml += '\n    <string name="'+subObj[options.string]+'">'+subObj[options.value]+"</string>";
    Logger.log("sub:"+subObj);
  }
  xml += "\n</resource>";
  return xml;
}

function json2strings(o, options) {
  var strings = "";
  for(var m in o){
    var subObj = o[m];
    strings += '"'+subObj[options.string]+'" = "'+subObj[options.value]+'"\n\n';
    Logger.log("sub:"+subObj);
  }
  return strings;
}
