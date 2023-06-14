function senderMail() {
    const id = 'ID GOES HERE'; // Update Sheets workbook id here
    const sheet = SpreadsheetApp.openById(id).getSheetByName('messages'); // worksheet must be named 'messages'
    const data = sheet.getDataRange().getValues();
    const users = data.slice(1);
    users.forEach(user =>{
      const email = user[2]; // emails from 3rd column
      const message = `Message Goes Here`;
      const subject = `Subject Goes Here`;
      GmailApp.sendEmail(email,subject,message);
    })
  }
  
