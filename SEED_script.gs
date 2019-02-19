var HEADER_LIST = [];
var ANDROID_FOLDER_NAME = "Android";
var ANDROID_FILE_NAME = "string.xml";
var iOS_FOLDER_NAME = "iOS";
var iOS_FILE_NAME = "Localizable.strings";

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{
        name: "Copy And Export",
        functionName: "copyAndExport"
    }];
    ss.addMenu("Localization", menuEntries);

}

function copyAndExport() {
    var html = HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setWidth(700)
        .setHeight(600);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
        .showModalDialog(html, 'Copy And Export');
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

//Save Usage
function saveResult(content, fileType) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var name = sheet.getName();
    // get the folder this spreadsheet locates in
    var thisFileId = SpreadsheetApp.getActive().getId();
    var thisFile = DriveApp.getFileById(thisFileId);
    var folder = thisFile.getParents().next();
    var d = new Date();
    var find = '';
    var re = new RegExp(find, 'g');
    if (fileType == "Android") {
        content = formatBeautifier(content, '&#60;', '<');
        content = formatBeautifier(content, '&#62;', '>');
        content = formatBeautifier(content, '<br>', '\n');

        var folders = folder.getFolders();
        while (folders.hasNext()) {
            var childFolder = folders.next();
            if (childFolder.getName() == ANDROID_FOLDER_NAME) {
                folder.removeFolder(childFolder);
            }
        };
        var subFolder = folder.createFolder(ANDROID_FOLDER_NAME);
        subFolder.createFile(ANDROID_FILE_NAME, content);
    } else {
        content = formatBeautifier(content, '<br>', '\n');

        var folders = folder.getFolders();
        while (folders.hasNext()) {
            var childFolder = folders.next();
            if (childFolder.getName() == iOS_FOLDER_NAME) {
                folder.removeFolder(childFolder);
            }
        }
        var subFolder = folder.createFolder(iOS_FOLDER_NAME);
        subFolder.createFile(iOS_FILE_NAME, content);
    }
    Logger.log("Content:" + content);
}

function formatBeautifier(source, target, replacement) {
    var re = new RegExp(target, 'g');
    return source.replace(re, replacement);
}

//COPY usage
function getData() {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var sheet = activeSpreadsheet.getActiveSheet()

    var activeRange = sheet.getDataRange();
    var data = activeRange.getValues();
    var array = [];
    for(var i = 0; i < data.length; i ++) {
       // Ignore first two columns for checking merge cell.
       var isMerge = sheet.getRange(i + 1, 3, 1, sheet.getMaxColumns()).isPartOfMerge();
       array.push({
        "isMerge": isMerge,
        "data": data[i]
       });
    }
    return array;
}

function getHeader() {
  
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var sheet = activeSpreadsheet.getActiveSheet()
    var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getMaxColumns());
    var headers = headersRange.getValues()[0];
    HEADER_LIST = normalizeHeaders_(headers);
    return (HEADER_LIST);
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

function dateFormat(d) {
    var fileName = d.getFullYear();
    if (d.getMonth() + 1 < 10) {
        fileName += '0' + (d.getMonth() + 1)
    } else {
        fileName += '' + (d.getMonth() + 1)
    }
    if (d.getDate() < 10) {
        fileName += '0' + d.getDate()
    } else {
        fileName += '' + d.getDate()
    }
    if (d.getHours() < 10) {
        fileName += '0' + d.getHours()
    } else {
        fileName += '' + d.getHours()
    }
    if (d.getMinutes() < 10) {
        fileName += '0' + d.getMinutes()
    } else {
        fileName += '' + d.getMinutes()
    }
    return fileName;
}

function highlightRow(row) {
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = activeSpreadsheet.getActiveSheet();
    var range = sheet.getRange(row + 1, 1, 1, sheet.getMaxColumns());
  
    range.setBackground("yellow");
}
