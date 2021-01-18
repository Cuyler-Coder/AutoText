  var sheet = SpreadsheetApp.getActiveSheet(); //Defines the variable sheet to contain whatever the active sheet is



function FindRows() {
range = SpreadsheetApp.getActiveSheet().getLastRow(); //Get the number of rows of texts in case new data is added or removed
return range-1; //subtracts 1 because the top row is only for labling
}

// Only run m-f business hours
function shouldRunTrigger() { 
  var days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var date = new Date();
  var day = days[date.getDay()];
  var hours = date.getHours();
  if ((day === "Fri" && hours >= 8 && hours <= 17) || (day === "Thu" && hours >= 8 && hours <= 17) || (day === "Wed" && hours >= 8 && hours <= 17) || (day === "Tue" && hours >= 8 && hours <= 17) ||(day === "Mon" && hours >= 8 && hours <= 17) ) {
    return true;
  }
  return false;
}

// Sends texts through email system. The phone number is defined the the cell A2
function sendEmails() {

  var startRow = 2; // First row of data to process
  var numRows = FindRows(); // Number of rows to process
  var indexPos = sheet.getRange(1, 3).getValue(); // Where the play head is. This head moves down the list until it reaches the bottom
  var odds = sheet.getRange(2,3).getValue(); // This is the % odds that the text will send every interval
  var oddsRate = sheet.getRange(2,4).getValue(); // The is the rate the odds increase every time a text does not get sent
  
  var range = sheet.getRange(2, 2, numRows); // The first two numbers are the starting cell (row and column) and the last number is the number of rows down from that point
  //Logger.log(range.getValues().reverse());
  
  //If the number of the play head position goes past the end of the list randomize the list and move the play head back to the starting position
    if(indexPos > numRows){
   indexPos = startRow; 
      checkArray(range.getValues());
    range.setValues(shuffleArray(range.getValues()));
  }

//sets the cell D1 to the value of numRows so that it is easier to see how many lines exist in the sheet
  sheet.getRange('d1').setValue(numRows);

// stops the code if it is outside business hours as difned the the function. I know this is ugly.
  if (!shouldRunTrigger()) return;
  
  // get a random number from 0 to 1 and compare it to the current odds. If it is a bigger number increase the odds by the odds rate and end the process. Also set the values for the odds and the odds rate cells so that it is easier to see from the spreadsheet.
  if(Math.random()>= odds){
    odds += oddsRate;
    oddsRate += .0002; // actual odds rate
    sheet.getRange(2, 3).setValue(odds);
    sheet.getRange(2, 4).setValue(oddsRate);
    
    return;
  }
  
    var emailAddress = sheet.getRange(2, 1).getValue(); // First column
    var message = sheet.getRange(indexPos, 2).getValue(); // Second column
    var timeDelay = Math.random(); //add a random delay to the message so that it does not send in predicable patterns.
    Utilities.sleep(timeDelay*300000);
    var subject = 'Sending emails from a Spreadsheet';
    MailApp.sendEmail(emailAddress, subject, message); //send the text using gmail
  indexPos += 1; // incrament the playhead position after sending the text
  sheet.getRange(2, 3).setValue(0) //reset the odds back to 0
  sheet.getRange(2, 4).setValue(oddsRate/10) //reduce the odds rate by an order of magnitude
  
  sheet.getRange(1, 3).setValue(indexPos); // Update the index position to reflect the playhead

  return;
}

// this is the shuffle function that is used to randomize the list
function shuffleArray(array) {
  var i, j, temp;
  for (i = array.length-1; i > 0; i--) {
    j = Math.floor(Math.random() * (i + 1));
    temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  Logger.log(array);
  return array;
}

// I don't remember what this does but it looks like a debugger I was using for something
function checkArray(ar) {
  if (ar.constructor !== Array) {
    Logger.log('Not an array'); return;
  }  
  var height = ar.length; 
  if (ar[0].constructor !== Array) {
    Logger.log('Not a double array'); return;
  }
  var width = ar[0].length;
  for (var i = 0; i < height; i++) {
    if (ar[i].length !== width) {
      Logger.log('Not rectangular: row '+(i+1)+' has length '+ar[i].length+' instead of '+width); 
      return;
    }
  }
  Logger.log('Rectangular of height '+height+' and width '+width);
}
