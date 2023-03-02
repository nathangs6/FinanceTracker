/*
Name: financeScripts.js
Author: Nathan Gurrin-Smith
Description: Contains the scripts used to support the Spending spreadsheet
*/

/*
Define the global constants.
Thanks to KrzFra (edited by Magne) for the idea
https://stackoverflow.com/questions/24721226/how-to-define-global-variable-in-google-apps-script
*/
function globalConstants() {
    var constants = {
      homeSheet: "Home",
      parameterSheet: "Parameters",
      expenseSheet: "Expenses",
      incomeSheet: "Income",
      analysisSheet: "Analysis",
      expenseRow: 6,
      numExpenseCols: 6,
      incomeRow: 10,
      numIncomeCols: 5
    };
    return constants;
  }
  
  // General Scripts
  
  /*
  getFirstEmptyRow finds the first empty row in a spreadsheet, only looking at column A
  Input:
    - spr: the spreadsheet to check
  Output: The number corresponding to the first empty row
  */
  function getFirstEmptyRow(spr) {
    var col = spr.getRange('A:A');
    var vals = col.getValues();
    var ct = 0;
    while (vals[ct] && vals[ct][0] != "") {
      ct++;
    }
    return (ct+1);
  }
  
  // Parameter Scripts
  
  function changeCategory() {
    // Initialize variables
    var gc = globalConstants();
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var parameterSheet = spr.getSheetByName(gc.parameterSheet);
    var expenseSheet = spr.getSheetByName(gc.expenseSheet);
  
    var oldCategory = parameterSheet.getRange(3,2).getValue();
    var newCategory = parameterSheet.getRange(3,3).getValue();
  
    if ((oldCategory === "") || (newCategory === "")) {
      throw "Error: old category or new category is empty!";
    }
  
    var endRow = getFirstEmptyRow(expenseSheet);
    
    for (var row=2;row<endRow;row++) {
      if (expenseSheet.getRange(row,4).getValue() === oldCategory) {
        expenseSheet.getRange(row,4).setValue(newCategory);
      }
    }
  
    parameterSheet.getRange(3,2,1,2).clearContent();
  }
  
  function changeSubCategory() {
    var gc = globalConstants();
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var parameterSheet = spr.getSheetByName(gc.parameterSheet);
    var expenseSheet = spr.getSheetByName(gc.expenseSheet);
  
    var newCategory = parameterSheet.getRange(7,2).getValue();
    var oldSubCategory = parameterSheet.getRange(7,3).getValue();
    var newSubCategory = parameterSheet.getRange(7,4).getValue();
  
    if ((oldSubCategory === "") || ((newSubCategory === "") && (newCategory === ""))) {
      throw "Error: old subcategory is empty or (new subcategory and new category) are empty!";
    }
  
    var endRow = getFirstEmptyRow(expenseSheet);
  
    for (var row=2;row<endRow;row++) {
      if (expenseSheet.getRange(row,5).getValue() == oldSubCategory) {
        if (newSubCategory != "") {
          expenseSheet.getRange(row,5).setValue(newSubCategory);
        }
        if (newCategory != "") {
          expenseSheet.getRange(row,4).setValue(newCategory);
        }
      }
    }
  
    parameterSheet.getRange(7,2,1,3).clearContent();
  }
  
  // Expense Scripts
  
  /*
  addExpense adds an expense entry to the expense data sheet
  */
  function addExpense() {
    // Setup variables
    var gc = globalConstants();
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var home = spr.getSheetByName(gc.homeSheet);
    var data = spr.getSheetByName(gc.expenseSheet)
    var row = getFirstEmptyRow(data);
    inputDataRange = home.getRange(gc.expenseRow,2,1,gc.numExpenseCols);
    outputDataRange = data.getRange(row,1,1,gc.numExpenseCols);
  
    // Do the work
    outputDataRange.setValues(inputDataRange.getValues());
    inputDataRange.clearContent();
  }
  
  // Income scripts
  
  /*
  getDestinationList takes in a profile string and spits out an array containing the profile contents
  Input:
    - profile: a string of the form 'Destination1,Percentage1,Destination2,Percentage2' etc.
  Output: an array of the form ['Destination1', Percentage1, 'Destination2', Percentage2] etc.
  Throws an error if the profile does not exist.
  */
  function getDestinationList(profile) {
    var gc = globalConstants();
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var home = spr.getSheetByName(gc.homeSheet);
    var parameters = spr.getSheetByName(gc.parameterSheet);
  
    // Find the destination list that corresponds to the selected profile
    if (profile == "Custom") {
      // if it's the custom profile, check the data input
      return home.getRange(10,6).getValue().split(',');
    } else {
      // if it's not the custom profile, loop through the pre-constructed profiles
      var profiles = parameters.getRange(12,5,3,1).getValues();
      for (var i=0;i<profiles.length;i++) {
        if (profile == profiles[i]) {
          // Found a profile! Compile the destination list
          return parameters.getRange(12+i,6).getValue().split(',');
        }
      }
    }
    throw ("Error: profile not found!");
  }
  
  
  /*
  constructProfileDict does two tasks:
    1. Takes in a destination list and spits out a dictionary mapping destinations to percentages
    2. Verifies the percentages add up to 100%
  Input:
    - destinationList: an array of the form ['Destination1', Percentage1, 'Destination2', Percentage2] etc.
  Output: a dictionary of the form {'Destination1': Percentage1, 'Destination2': Percentage2} etc.
  Throws an error if the total percentage is not 100.
  */
  function constructProfileDict(destinationList) {
    var profileDict = {
    };
    var s = 0;
    for (var i=0;i<destinationList.length;i+=2) {
      profileDict[destinationList[i]] = +destinationList[i+1];
      s += +destinationList[i+1];
    }
    // Verify that the total percentage is 100
    if (s != 100) {
      throw ("Error: Total percentage not 100!");
    }
    return profileDict;
  }
  
  /*
  addIncome adds an income entry to the income data sheet.
  */
  function addIncome() {
    // get sheet data
    var gc = globalConstants();
    var spr = SpreadsheetApp.getActiveSpreadsheet();
    var home = spr.getSheetByName(gc.homeSheet);
    var data = spr.getSheetByName(gc.incomeSheet)
    var row = getFirstEmptyRow(data);
  
    // get input values
    var values = [[
      home.getRange(gc.incomeRow,2).getValue(), // source
      "",
      0,
      home.getRange(gc.incomeRow,4).getValue(), // date
      home.getRange(gc.incomeRow,7).getValue()  // notes
    ]];
    var amount = home.getRange(gc.incomeRow,3).getValue();
  
    // get profile and compile its saving breakdown
    var profile = home.getRange(10,5).getValue();
    var destination = constructProfileDict(getDestinationList(profile));
  
    // Fill in income data sheet
    var i = 0;
    for (const saving in destination) {
      // update values to have the correct saving title and amount
      values[0][1] = saving;
      values[0][2] = amount * (destination[saving]/100);
      // update the output range
      data.getRange(row+i,1,1,gc.numIncomeCols).setValues(values);
      i = i + 1;
    }
    home.getRange(gc.incomeRow,2,1,6).clearContent()
  }
  