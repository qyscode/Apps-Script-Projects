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
