/*
  Instructions:

  1:  Name one of your tabs "Addresses"
  2:  Make sure your columns are in this order:
        Name,
        Street,
        Street2,
        City,
        State,
        Zip
  3:  In the top menu, select Tools > Script Editor
  4:  Delete everything inside the code editor (the empty function thing)
  5:  Copy and paste this code into the code editor and click the "Save" icon
  6:  Refresh the spreadsheet (the Script Editor will probably close - that's ok. You're done with it)
  7:  There will be a new button at the top of the page called "Custom Scripts"
  8:  Select Custom Scripts > Format Emails
  9:  Grant permissions when it prompts you to. (You'll only do this once per Google account)
  10: Run the script one more time if it does not run the first time
  
  Hopefully this helps, brother! This is just a first pass. I'm going to keep building this out.
  Love ya, man!  
*/ 

const ss = SpreadsheetApp.getActiveSpreadsheet();
const input = ss.getSheetByName("Addresses");

// Create menu items

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Scripts')
    .addItem('Format Addresses','main')
    .addToUi();
}

// Create main function

function main() {
  checkForOuput();
  formatAddresses();
  prettify();
  SpreadsheetApp.getUi().alert(
    "Behold the Magic...", 
    "Addresses formatted and transferred to new sheet", 
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// Create ouput sheet if not in spreadsheet

function checkForOuput() {
  var search = ss.getSheetByName("Formatted Addresses");
  if (!search) {
   ss.insertSheet("Formatted Addresses");
  }
}

// Create function to transpose addresses

function formatAddresses() {
  var output = ss.getSheetByName("Formatted Addresses");
  var range = input.getDataRange();
  var values = range.getValues();
  var destRange = output.getDataRange();
  destRange.clear();
  var finalRange = output.getRange(1,1,values[0].length,values.length)
  finalRange.setValues(Object.keys(values[0]).map ( function (columnNumber) {
    return values.map( function (row) {
      return row[columnNumber];
    });
  })); 
}

// Create function to set sheet formatting

function prettify() {
  var output = ss.getSheetByName("Formatted Addresses");
  var destRange = output.getDataRange();
  var destValues = destRange.getValues();
  destRange.setNumberFormat('@STRING@');
  output.autoResizeColumns(1, destValues[0].length);
  destRange.setHorizontalAlignment("center");
}