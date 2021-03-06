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

  // Now, grab the number of people staying   
  var numPeople = grabColumnValues(2,2);
  
  
  var emailsList = "mjsgreen01@gmail.com,dmj133@juno.com,sdmjgreen@gmail.com,kris_aoay@yahoo.com,jsw.weinberg@gmail.com";
  var emailSubject = "Green House Rental Reminder";
  var warning_count = 0;
  var msg = "";
  var message = "";
  var days_left_message = "";


  /**
  * Normalize the dates
  * Starts out in format "Mon Mar 21 2016 00:00:00 GMT-0400 (EDT)" (but as a Date object)
  * Should end as "Mon Mar 21 2016"
  */
  var normalizeDate = function (date) {
    if(date){
      // get the cut-off point 
      var stringEndIndex = date.toString().indexOf(" 00:00:00");
      var normlizedDate = date.toString().slice(0, stringEndIndex);
    }

    return normlizedDate;
  };

  /**
  * Construct email-body message based on number of days left
  * first param is T/F - if F, the person is moving out, not in
  */
  var constructMessage = function(movingIn, reminder_name, days_left, date, move_out_date, peopleInParty){
    // normalize the date strings
    date = normalizeDate(date);
    move_out_date = normalizeDate(move_out_date);

    // set dynamic text in message
    if(days_left === 1){
      days_left_message = "tomorrow";
    }else if(days_left === 7){
      days_left_message = "in one week";
    }else if(days_left === 14){
      days_left_message = "in two weeks";
    }else if(days_left === 30){
      days_left_message = "in one month";
    }else if(days_left === 60){
      days_left_message = "in two months";
    }else{
      days_left_message = "in "+days_left+" days";
    }

    // build the message
    if(movingIn){
      message = "Reminder: "+reminder_name+" will arrive at the Green's house as a party of "+peopleInParty+" "+days_left_message+", on "+date+", and will check out on "+move_out_date+"\n ";
    }else{
      message = "Reminder: "+reminder_name+" will check OUT of the Green's house "+days_left_message+", on "+date+".\n";
    }
    return message;
  };

  /**
  * Check if move-in reminder should be sent, and send it
  */
  var checkAndSendMovein = function(daysLeft){
    warning_count = 0;
    // Loop over the days left till movein values
    for (var i = 0; i <= numRows-1; i++) {
      var days_left = days_left_movein[i][0];
      if(days_left == daysLeft) {
        // if it's exactly 5, do something with the data.
        var reminder_name = guestNames[i][0];
        var moveIn = moveInDate[i][0];
        var moveOut = moveOutDate[i][0];
        var peopleInParty = numPeople[i][0];
        
        msg = constructMessage(true, reminder_name, days_left, moveIn, moveOut, peopleInParty);
        warning_count++;
      }
    }
    
    //send the email if specified # of days are left
    if(warning_count) {
      MailApp.sendEmail(emailsList, emailSubject, msg);
    }

  };


  /**
  * Check if move-out reminder should be sent, and send it
  */
  var checkAndSendMoveout = function(daysLeft){
    warning_count = 0;
    // Loop over the days left till moveout values
    for (var i = 0; i <= numRows-1; i++) {
      var days_left_out = days_left_moveout[i][0];
      if(days_left_out == daysLeft) {
        // if it's exactly 3, do something with the data.
        var reminder_name = guestNames[i][0];
        var moveOut = moveOutDate[i][0];
        
        msg = constructMessage(false, reminder_name, days_left_out, moveOut);
        warning_count++;
      }
    }
    
    //send the email if specified # of days are left
    if(warning_count) {
      MailApp.sendEmail(emailsList, emailSubject, msg);
    }

  };
  


  checkAndSendMovein(1);
  
  
  checkAndSendMovein(7);
  
  
  checkAndSendMovein(14);
  

  checkAndSendMovein(30);


  checkAndSendMovein(60);
  
  
  

  checkAndSendMoveout(3);
  

  checkAndSendMoveout(1);
  
  
  
}

