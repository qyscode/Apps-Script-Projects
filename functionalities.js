// Log run time of program
let runStart = new Date();
let runEnd = new Date();
let runTime = runEnd - runStart;
Logger.log("Excel time: " + runTime + "ms");

// Duplicate a worksheet within a Sheet
const ss = SpreadsheetApp.openById("ss-id") // spreadsheet
const ws = ss.getSheetByName("ws-name"); // worksheet
const os = ws.copyTo(ss).setName("os-name") // output sheet

// Search for keyword among range to last row
const sourceSheet = ss.getSheetByName(ws-name);
const relevantRange = sourceSheet.getRange(`A5:A`+ sourceSheet.getLastRow());
const relevantCells = relevantRange.createTextFinder("keyword").findAll(); // Returns a Range[] data type

// Dates
// let dateFormat = `${day}/${month}/${year}`;
const date = new Date();
let day = double_Digit_(date.getDate());
let month = double_Digit_(date.getMonth() + 1); // JS counts month from 0
let year = date.getFullYear();
function double_Digit_(number) {
    if (number < 10) {
        number = "0" + number;
        return number
    }
    return number
}

// Misc
// `${worksheetName.slice(10)}`
