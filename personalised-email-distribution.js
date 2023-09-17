function senderMail() {
    const sheet = SpreadsheetApp.openById(id).getSheetByName('ws-name');
    const data = sheet.getDataRange().getValues();
    const users = data.slice(1);
    users.forEach(user =>{
      const email = user[2]; // emails from 3rd column
      GmailApp.sendEmail(email,subject,message);
    })
  }
