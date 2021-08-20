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
  var values = input.getDataRange().getValues();
  var name = values[0].indexOf("NAME");
  var street1 = values[0].indexOf("Street");
  var street2 =  values[0].indexOf("Street #2");
  var city = values[0].indexOf("CITY");
  var state = values[0].indexOf("STATE");
  var zip = values[0].indexOf("ZIP");
  var destRange = output.getDataRange();
  destRange.clear();
  var nameArr = [];
  var stArr = [];
  var cityArr = [];
  var stackedAddress = [];

  values.filter (function (row) {
    return (row[name] !== "");
  })
  .forEach(function (row) {
    nameArr.push(row[name]);
    if(row[street2] !== "") {
      stArr.push(row[street1] + " " + row[street2]);
    }
    else {
      stArr.push(row[street1]);
    };
    cityArr.push(row[city] + ", " + row[state] + " " + row[zip]);
  });
  stackedAddress.push([nameArr],[stArr], [cityArr]);
  var nameRange = output.getRange(1, 1, 1, stackedAddress[0][0].length);
  var stRange = output.getRange(2, 1, 1, stackedAddress[0][0].length);
  var cityRange = output.getRange(3, 1, 1, stackedAddress[0][0].length);
  nameRange.setValues(stackedAddress[0]);
  stRange.setValues(stackedAddress[1]);
  cityRange.setValues(stackedAddress[2]);
}

// Create function to set sheet formatting

function prettify() {
  var output = ss.getSheetByName("Formatted Addresses");
  var destRange = output.getDataRange();
  var destValues = destRange.getValues();
  destRange.setNumberFormat('@STRING@');
  output.autoResizeColumns(1, destValues[0].length);
  destRange.setHorizontalAlignment("center");
  output.setFrozenColumns(1);
}