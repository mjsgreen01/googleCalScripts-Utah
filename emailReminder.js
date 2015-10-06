/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Read Data",
    functionName : "readRows"
  },{
    name : "Check Reminders",
    functionName : "checkReminder"
  }];
  spreadsheet.addMenu("Script Center Menu", entries);
}






function checkReminder() {
  // get the spreadsheet object
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // set the first sheet as active
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]);
  // fetch this sheet
  var sheet = spreadsheet.getActiveSheet();
  
  // figure out what the last row is
  var lastRow = sheet.getLastRow();

  // the rows are indexed starting at 1, and the first row
  // is the headers, so start with row 2
  var startRow = 2;

  // get the number of rows in the spreadsheet
  var range = sheet.getRange(2,5,lastRow-startRow+1,1 );
  var numRows = range.getNumRows();

  var grabColumnValues = function(row, column){
    // args are: (startRow, startCol, numRows, numCols)
    var range = sheet.getRange(row, column, lastRow-startRow+1, 1 );
    return range.getValues();
  };

  // grab column 5 (the 'days left' till movein column) 
  var days_left_movein = grabColumnValues(2,5);
  
  // grab column 6 (the 'days left' till moveout column) 
  var days_left_moveout = grabColumnValues(2,6);
  
  // grab column 3 (the 'moveInDate' column) 
  var moveInDate = grabColumnValues(2,3);
  
  // grab column 4 (the 'moveOutDate' column) 
  var moveOutDate = grabColumnValues(2,4);
  
  // Now, grab the reminder name column
  var guestNames = grabColumnValues(2,1);
  
  
  var emailsList = "mjsgreen01@gmail.com,dmj133@juno.com,sdmjgreen@gmail.com,kris_aoay@yahoo.com,jsw.weinberg@gmail.com";
  var emailSubject = "Green House Rental Reminder";
  var warning_count = 0;
  var warning_count2 = 0;
  var warning_count3 = 0;
  var warning_count4 = 0;
  var warning_count5 = 0;
  var warning_count6 = 0;
  var warning_count7 = 0;
  var msg = "";
  var msg2 = "";
  var msg3 = "";
  var msg4 = "";
  var msg5 = "";
  var msg6 = "";
  var msg7 = "";
  var days_left_message = "";

  var constructMessage = function(reminder_name, days_left, moveIn){
    if(days_left === 1){
      days_left_message = "tomorrow";
    }else if(days_left === 14){
      days_left_message = "in two weeks";
    }else if(days_left === 30){
      days_left_message = "in one month";
    }else if(days_left === 60){
      days_left_message = "in two months";
    }else{
      days_left_message = "in "+days_left+" days";
    }

    var message = "Reminder: "+reminder_name+" will arrive at the  Green's house "+days_left_message+", on "+moveIn+".\n";
    return message;
  };

  var checkAndSendMovein = function(daysLeft){
    var warning_count = 0;
    // Loop over the days left till movein values
    for (var i = 0; i <= numRows-1; i++) {
      var days_left = days_left_movein[i][0];
      if(days_left == daysLeft) {
        // if it's exactly 5, do something with the data.
        var reminder_name = guestNames[i][0];
        var moveIn = moveInDate[i][0];
        
        msg = constructMessage(reminder_name, days_left, moveIn);
        warning_count++;
      }
    }
    
    //send the email if specified # of days are left
    if(warning_count) {
      MailApp.sendEmail(emailsList, 
          emailSubject, msg);
    }

  };
  


  checkAndSendMovein(1);
  
  
  checkAndSendMovein(5);
  
  
  checkAndSendMovein(14);
  

  checkAndSendMovein(30);


  checkAndSendMovein(60);
  
  
  
  
  
  // Loop over the days left till moveout values
  for (var i = 0; i <= numRows-1; i++) {
    var days_left_out = days_left_moveout[i][0];
    if(days_left_out == 3) {
      // if it's exactly 3, do something with the data.
      var reminder_name = guestNames[i][0];
      var moveOut = moveOutDate[i][0];
      
      msg2 = msg2 + "Reminder: "+reminder_name+" will check out of the Green's house in "+days_left_out+" days, on "+moveOut+".\n";
      warning_count2++;
    }
  }
  
  //send the email if specified # of days are left
  if(warning_count2) {
    MailApp.sendEmail(emailsList, 
        emailSubject, msg2);
  }
  
  
  
  
  
   // Loop over the days left till moveout values
  for (var i = 0; i <= numRows-1; i++) {
    var days_left_out = days_left_moveout[i][0];
    if(days_left_out == 1) {
      // if it's exactly 1, do something with the data.
      var reminder_name = guestNames[i][0];
      var moveOut = moveOutDate[i][0];
      
      msg7 = msg7 + "Reminder: "+reminder_name+" will check OUT of the Green's house tomorrow, on "+moveOut+".\n";
      warning_count7++;
    }
  }
  
  //send the email if specified # of days are left
  if(warning_count7) {
    MailApp.sendEmail(emailsList, 
        emailSubject, msg7);
  }



  
  
};