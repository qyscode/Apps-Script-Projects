let compilation_start = new Date();

// Return the current date in currentDate
const date = new Date();
let day = date.getDate();
if (day < 10) {
  day = "0" + day;
  }
let month = date.getMonth() + 1;
if (month < 10) {
  month = "0" + month;
}
let year = date.getFullYear();

let currentDate = `${day}/${month}/${year}`;
var folderDateName = `${year}_${month}_${day}`;

function sendEmail() {

  // Define email parameters
  var recipient = "RECIPIENT HERE";
  var subject = `Report (${currentDate})`;
  var body = "Hi all,\n\nAttached are the reports for today.\n\nRegards,\nName\n\n\n";

  // Extract Excel file
  var fileId = "GOOGLE SHEETS FILE ID HERE";
  var excel_file = UrlFetchApp.fetch(
    "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + fileId + "&exportFormat=xlsx",
    {
      "headers": {Authorization: "Bearer " + ScriptApp.getOAuthToken()},
      "muteHttpExceptions": true
    }
  ).getBlob().setName(`Excel Report (${currentDate})` + ".xlsx")

  // Extract ZIP/PDF
  let zip_file_name = 'Reports for '+folderDateName+'.zip'
  // This line below searches for a folder with the format YYYY_MM_DD within Drive. For a specific, fixed folder, getFolderById may be preferred.
  const folderIterator = DriveApp.getRootFolder().getFoldersByName(folderDateName);
  var folder = folderIterator.next();

  var zipped = Utilities.zip(getBlobs(folder, ''), zip_file_name);
  var zip_id = folder.createFile(zipped).getId();
  var zip_file = DriveApp.getFileById(zip_id);

  function getBlobs(rootFolder, path) {
    var blobs = [];
    var files = rootFolder.getFiles();
    while (files.hasNext()) {
      var file = files.next().getBlob();
      file.setName(path+file.getName());
      blobs.push(file);
    }
    return blobs;
  }

// Send email
MailApp.sendEmail(recipient, subject, body, {name:"OPTIONAL CUSTOM SENDER NAME",attachments: [excel_file, zip_file]})}

let compilation_end = new Date();
let time = compilation_end - compilation_start;
Logger.log("compilation time: " + time + "ms");

function sortANDrename() {
  let renamingstart = new Date();

  // correct workbook name "Daily Report | MTH 2023"
  let actual_numeric_month = date.getMonth();
  const monthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const actual_month_name_three = monthNames[actual_numeric_month];
  const workbookName = `Daily Report | ${actual_month_name_three} 2023`;
  console.log(`Sheet name is: ${workbookName}`);
  const sheet = SpreadsheetApp.openById(fileId).getSheetByName('workbookName');


  // SORT FROM first S/N of the day
  var firstSNoftheday = `${folderDateName}_1`
  var textFinderSORT = sheet.createTextFinder(firstSNoftheday).findNext(); // Finds first instance of file name in sheet
  var rowSORT = textFinderSORT.getRow();
  //var colSORT = textFinderSORT.getColumn(); // Column is not needed
  var range = sheet.getRange(`A${rowSORT}:W`);
  range.sort({column: 2, ascending: true}); // Sorts "range" by column B

  // For each file in the folder, get the name, search for the name in the excel sheet, look at the TITLE on the same row as the matching serial number, rename the file to that title.
  const targetFolderPDF = DriveApp.getFolderById("ID HERE"); // Folder containing files
  const folderIterator = targetFolderPDF.getFoldersByName(folderDateName); // FIND CURRENT DATE FOLDER BASED ON NAME
  var file_iterator = folderIterator.next().getFiles(); // 
  while (file_iterator.hasNext()) {
    const pdfFile = file_iterator.next();  // SELECTS THE NEXT FILE
    var pdfFileName = pdfFile.getName();
    var textFinder = sheet.createTextFinder(pdfFileName).findNext(); // Finds first instance of PDF file name in sheet
    var row = textFinder.getRow();
    //var col = textFinder.getColumn();
    var title_of_pdf = sheet.getRange(int,int,int,int).getValues();
    pdfFile.setName(title_of_pdf)
  }

  let renamingend = new Date();
  let renamingtime = renamingend - renamingstart;
  Logger.log("Sort and renaming time: " + renamingtime + "ms");
}

