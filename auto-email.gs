function sendEmails() {  
  var timeStamp = Utilities.formatDate(new Date(), "GMT+8", "MM/dd/yyyy H:mm:ss");
  var dateToday = new Date();
  var sheet = SpreadsheetApp.getActive().getSheetByName("F&L Wishlist");
  var startRow = 2; // First row of data to process
  var numRows = 2; // Number of rows to process
  var dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  var data = dataRange.getValues();                         
  
  var html = DriveApp.getFilesByName('Brand Acquisition Auto-Mailer V9.html').next().getBlob().getDataAsString();

// cofigurations for sending once every 2 weeks
  var secondRow = data[0];
  var lastSent = secondRow[9];
  var dateSent = new Date(lastSent);
  var dateDiff = Math.floor((dateToday.getTime()-dateSent.getTime())/(24*3600*1000));
  Logger.log('dateToday: '+dateToday);
  Logger.log('dateSent: '+dateSent);
  Logger.log('Date Diff: '+dateDiff);
  Logger.log(dateDiff);
  if(dateDiff > 13) {
  Logger.log('sending emails');
    for (var i = 0; i < data.length; ++i) {
      var row = data[i];
      var emailSent = row[0];
      var brand = row[2];
      var status = row[5];
      var emailAddress1 = row[6];
      var emailAddress2 = row[7];
      var emailAddress3 = row[8];
            
      var name = 'Shopee Malaysia'
      var subject = 'Start selling with Shopee Malaysia today!';
      var body = html
      
      Logger.log(status);
      if(status == 'On Hold' || status == 'Rejected' || status == 'Engaging') {
        if(emailAddress1 !== "") {      
          tryCatchEmail(emailAddress1,name,subject,body);
        }
        
        if(emailAddress2 !== "") {      
          tryCatchEmail(emailAddress2,name,subject,body);
        }
        
        if(emailAddress3 !== "") {      
          tryCatchEmail(emailAddress3,name,subject,body);
        }
        
        if(emailAddress1 !== "" || emailAddress2 !== "" || emailAddress3 !== "") {
          sheet.getRange(startRow + i, 1).setValue(timeStamp);
          SpreadsheetApp.flush();
        }
       
        Logger.log('Sent Emails');
      }
            
    }
  } else {
    Logger.log('Chill, Not 2 weeks yet')
  }
}

// ================ //
// HELPER FUNCTIONS //
// ================ //

function tryCatchEmail(email,name,subject,body){
  try{
    MailApp.sendEmail({
      to: email,
      name: name,
      subject: subject,
      htmlBody: body
    });
  }
  catch(err){
    Logger.log(err);
  }
}

function properCase(phrase) {
    var regFirstLetter = /\b(\w)/g;
    var regOtherLetters = /\B(\w)/g;
    function capitalize(firstLetters) {
      return firstLetters.toUpperCase();
    }
    function lowercase(otherLetters) {
      return otherLetters.toLowerCase();
    }
    var capitalized = phrase.replace(regFirstLetter, capitalize);
    var proper = capitalized.replace(regOtherLetters, lowercase);    
    return proper;
}