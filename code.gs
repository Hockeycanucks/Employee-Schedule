// these are labels for the days of the week
var cal_days_labels = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

var cal_work_days = ['su', 'mo', 'tu', 'we', 'tr', 'fr', 'sa'];

// these are human-readable month name labels, in order
var cal_months_labels = ['January', 'February', 'March', 'April',
                     'May', 'June', 'July', 'August', 'September',
                     'October', 'November', 'December'];

// these are the days of the week for each month, in order
var cal_days_in_month = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

// The Employee object that stores employee information
function Employee(ss, name, workDays, hours) {
  this.ss = ss;
  this.name = name;
  this.workDays = workDays;
  this.hours = hours;
}

Employee.prototype = {
  constructor: Employee,
  
  getName: function() {
    return this.name;
  },
  
  checkWorkDay: function(dayNumber) {
    return this.workDays[dayNumber];
  },
  
  /**
  * Set the full working week for an employee.
  *
  * @param {Array} workDays An array of booleans to set the work week.
  */
  setWorkDays: function(workDays) {
    this.workDays = workDays;
  },
  
  setWorkDay: function(workDay, working) {
    this.workDays[workDay] = working;
  },
  
  addHours: function(amt) {
    this.hours += amt;
  }
};


/**
* A special function thats run when the spreadsheet is open
*
*/
function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('Schedule Options')
  .addItem('New Month', 'doGet')
  .addSeparator()
  .addItem('Add Pay Period', 'addPayPeriod')
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Remove Pay Period')
              .addItem('Pay Period 1', 'rePPOne')
              .addItem('Pay Period 2', 'rePPTwo')
              .addItem('Pay Period 3', 'rePPThree'))
  .addSeparator()
  .addItem('Add New Employee', 'addEmployee')
  .addSeparator()
  .addItem('Holiday Time Checker', 'holDoGet')
  .addSeparator()
  .addItem('Help', 'showSidebar')
  .addToUi();
}

function doGet() {
  //Open a dialog
  var htmlDlg = HtmlService.createHtmlOutputFromFile('Index')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(200)
      .setHeight(80);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'Enter the Month and Year');
}

function holDoGet() {
  var htmlDlg = HtmlService.createHtmlOutputFromFile('holiday')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(125);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'Calculate Holiday Pay');
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Help')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

/**
* A function that creates the calendar.
*
*/
function createCal(startDate) {
  // Check if the format on the settings page is done correctly.
  if(SpreadsheetApp.getActive().getSheetByName('Settings').getRange(2, 6).getValue() != 'Good') {
    throw Browser.msgBox('The format for the work days (column B) is incorrect. Please fix any errors and retry');
  }
  if(SpreadsheetApp.getActive().getSheetByName('Settings').getRange(3, 6).getValue() != 'Good') {
    throw Browser.msgBox('The format for the store hours (column D) is incorrect. Please fix any errors and retry');
  }
  
  // Check if there is already a sheet created.
  if (SpreadsheetApp.getActive().getSheetByName(startDate) != null) {
    throw Browser.msgBox('The calendar for ' + startDate + ' already exists.');
  } else {
    createSheet_(startDate);
  }
  
  for(var i=1; i<=14; i++) SpreadsheetApp.getActiveSheet().setColumnWidth(i, 85);
  
  var month = startDate.split(" ");
  addDayNumbers(findDay_(startDate), month[0], month[1]);
  
  addBorders_(startDate);
  
  populateEmployees(startDate);
  
  addHours(startDate);
  
  employeeVal(startDate);
  
  ppTemplate(startDate);
  
  holPayTemp(startDate);
}

function addPayPeriod(){
  var htmlDlg = HtmlService.createHtmlOutputFromFile('pp')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400)
      .setHeight(125);
  SpreadsheetApp.getUi()
      .showModalDialog(htmlDlg, 'New Pay Period');
}

function createEmployeeList(startDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var month = ss.getSheetByName(startDate);
  var setting = ss.getSheetByName("Settings");
  
  var employeeList = [];
  var row = 2;
  
  //Checks over the settings page for employee names.
  while (setting.getRange(row, 1).getValue().length != 0) {
    
    //Set up the variables an employee object uses. 
    var name = setting.getRange(row, 1).getValue();
    var workDays = [false,false,false,false,false,false,false];
    var workDaysSet = setting.getRange(row, 2).getValue().split(',');
    
    // Test if there is an issue with the work days.
    
    //Find which days an employee works.
    var j = 0;
    for (var i=0; i<workDays.length; i++) {
      if (workDaysSet[j] == cal_work_days[i]) {
        workDays[i] = true;
        j++;
      }
    }
    
    //Create a new employee and add them into the employee list.
    var employee = new Employee(ss, name, workDays, 0);
    employeeList.push(employee);
 
    row++;
  }
  
  return employeeList;
}

/**
* Adds the employee's names to calendar on the correct days
*
* @param {String} startDate The name of the sheet that needs to be populated.
*/
function populateEmployees(startDate) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var month = ss.getSheetByName(startDate);
  var setting = ss.getSheetByName("Settings");
  var cal = getCalValues(startDate);
      
  var employeeList = createEmployeeList(startDate);
  
  var lastDay = monthLastDay(startDate);
  
  var locked = false;
  var row = 2;
  //var startRow = row;
  var col = 0;
  while (!locked) {
    if (cal[row-1][col] == lastDay) locked = true;
    
    var rowReset = row;
    for (var i=0; i<employeeList.length; i++) {
      if (employeeList[i].checkWorkDay(col/2) && cal[rowReset-1][col] != '') {
        cal[row][col] = employeeList[i].name;
        row++;
      }
    }
    row = rowReset;
    
    //Restart the rows and columns to propper position.
    col += 2;
    if (col > 13) {
      col = 0;
      row += 10;
    }
  }
  month.setActiveRange(getCalRange(startDate)).setValues(cal);
}

/**
* Find the starting day of the week for a month.
*
* @param {String} startDay The Name of the calendar page.
* @return {Number} The number from 0-6 refering to a day of the week.
*/
function findDay_(startDay) {
  
  var strDate = startDay.split(" ");
  var monthInt = 0;
  
  //Find the month and covert to int.
  for (var i=0; i<cal_months_labels.length; i++) {
    if (strDate[0].toLowerCase() == cal_months_labels[i].toLowerCase()) {
      monthInt = i;
    }
  }
 
  // Create a new date object to find what the starting day of the week is for a month.
  var date = new Date(strDate[1], monthInt, 1, 0, 0, 0 ,0);
  var day = date.getDay();
  
  return day;
}

/**
* A function to add day numbers to the calendar.
*
* @param {Number} startDay The day the month starts on.
*/
function addDayNumbers(startDay, month, year) {
  
  var s = SpreadsheetApp.getActiveSheet();
  var range = s.getRange(2, 1);
  var endDay = 1;
  
  //Find the months last day.
  for (var i=0; i<cal_months_labels.length; i++) {
    if (cal_months_labels[i].toLowerCase() == month.toLowerCase()) {
      endDay = cal_days_in_month[i];
    }
  }
  
  // Check if it is a leap year.
  if (endDay == 28) {
    if ((year % 4 == 0 && year % 100 != 0) || year % 400 == 0){
      endDay = 29;
    }
  }
  
  // Add the day numbers to the calendar.
  var row = 2;
  for (var i=1; i<=endDay; i++) {
    
    s.getRange(row, startDay % 7 * 2 +1).setValue(i);
    
    startDay++;
    
    if (startDay%7 ==  0) {
      row += 10;
    }
  }
}

/**
* A function that creates the calendar template.
*
* @param {String} startDate The month and year of the sheet.
*/
function createSheet_(startDate) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet(startDate);
  var newSheet = SpreadsheetApp.getActiveSheet();
  
  //Set top bar with a grey background
  newSheet.getRange('A1:N1').setBackground('#336600').setFontWeight('bold').setHorizontalAlignment('center').setFontColor('white');
 
  //Merge the top bar and make columns smaller
  //Set borders and add days of the week
  for (var i=1; i<15; i+=1) {
    if (i%2 == 1) {
      newSheet.getRange(1, i, 1, 2).merge().setBorder(true, true, true, true, null, null);
      newSheet.getRange(1, i).setValue(cal_days_labels[Math.floor(i/2)]);
    }
    newSheet.setColumnWidth(i, 75);
  } 
}

/**
* Adds borders to calendar days.
*
* @param {String} startDate The month and year of the sheet.
*/
function addBorders_(startDate) {
  var s = SpreadsheetApp.getActive().getSheetByName(startDate);
  
  var endBorder = 42
  if (s.getRange(52, 1).getValue() != 0) {
    endBorder = 52;
  }
  
  //Set Borders for calendar days
  for (var i=2; i<=endBorder; i+=10) {
    for(var j=1; j<15; j+=2) {
      s.getRange(i, j, 1, 2).merge().setFontWeight('bold').setBackground('#e6e6e6');
      s.getRange(i, j, 10, 2).setBorder(true, true, true, true, null, null);
    }
  }
}

/**
* A function that gets the store hours from the Settings page.
*
* @return {String[]} Returns an array containing the store hours.
*/
function getHours_() {
  var ss = SpreadsheetApp.getActive();
  var setting = ss.getSheetByName('Settings');
  
  var hours = [7];
  for(var i=0; i<7; i++) {
    hours[i] = String(setting.getRange(i+2, 4).getValue());
  }
  return hours;
}

/**
* Adds the store hours to the employees.
*
* @parma {String} startDate The calendar to add hours to.
*/
function addHours(startDate) {
  var s = SpreadsheetApp.getActive().getSheetByName(startDate);
  var hours = getHours_();
  var cal = getCalValues(startDate);
  var calBox = Math.floor(cal.length / 10) * 7;
  
  var row = 3,
      col = 2;
  for (var i=0; i<calBox; i++) {
    // Set the correct format for a calendar's store hours.
    // Probably should use getNumberFormat() here.
    s.getRange(row, col, 9).setNumberFormats([['@STRING@'],['@STRING@'],['@STRING@'],['@STRING@'],['@STRING@'],['@STRING@'],['@STRING@'],['@STRING@'],['@STRING@']]);
    col += 2;
    if (col > 13) {
      col = 2;
      row += 10;
    }
  }
  
  var lastDay = monthLastDay(startDate);
  row = 2;
  col = 1;
  var locked = false;
  // Run until the last day is hit.
  while(!locked) {
    if(cal[row-1][col-1] == lastDay) locked = true;
    
    var inRow = row;
    while (cal[inRow][col-1] != "") {
      cal[inRow][col] = hours[Math.floor(col/2)];
      inRow++;
    }
    
    col += 2;
    if (col > 13) {
      col = 1;
      row += 10;
    }
  }
  
  s.setActiveRange(getCalRange(startDate)).setValues(cal);
}

/**
* Turn a date into a string.
*
* @param {Date} date A Date.
* @return {String} The date as a String. Format 'month year'.
*/
function getSheetDate(date) {
  var month = date.getMonth(),
      year = date.getFullYear();
  
  var sheetDate = (String(cal_months_labels[month]) + ' ' + String(year));
  
  return sheetDate;
}

/**
* Calculate the hours an employee has worked in a pay period.
*
* @param {String} name The name of the employee you want to find the hours of.
* @param {String} startPay The start date of the pay peiod.
* @param {String} endPay The end of the pay period.
* @param {String[][]} startRange The starting calendar.
* @param {String[][]} endRange The range of the ending calendar.
* @return {Number} The amount of hours the employee has worked.
* @customfunction
*/
function calcPayFormula(name, startPay, endPay, startRange, endRange) {
  var ss = SpreadsheetApp.getActive()
  
  // Date Object for the start and ending dates.
  var startDate = new Date(startPay),
      endDate = new Date(endPay);
  
  // Use the dates to find the start and end sheet names.
  var sSheetName = getSheetDate(startDate),
      eSheetName = getSheetDate(endDate);
  
  // Get the two sheets that need to be processed.
  var sSheet = ss.getSheetByName(sSheetName),
      eSheet = ss.getSheetByName(eSheetName),
      curSheet = sSheet;

  // Get the values of the start and end calendars
  var startCal = startRange,
      endCal = endRange,
      curCal = startCal;
  
  // Get the list of employees
  var emList = createEmployeeList();
  
  // Find the last day of the starting sheet.
  // Only needed if the pay period is over two months.
  var lastDay = monthLastDay(sSheetName);
  
  // Get the position of the emnployee in the employee list.
  var nameId = 0;
  for(var i=0; i<emList.length; i++) {
    if(emList[i].name == name) {
      nameId = i;
      break;
    }
  }
  
  var row = 2,
      col = 0,
      rowReset = row;
  var locked = false,
      inPay = false;
  while (!locked) {
    rowReset = row;
    // End when the last day has been hit.
    if (curCal[row-1][col] == endDate.getDate() && curSheet.getName() == eSheet.getName()) locked = true;
    
    if (startCal[row-1][col] == startDate.getDate()) {
      inPay = true;
    }
    
    // Only check hours while in the pay period.
    if(inPay) {
      while (row < rowReset + 8) {
        if (curCal[row][col] == name && curCal[row][col+1] != '') {
          var hoursString = curCal[row][col+1].split('-');
          
          // Error check the hoursString.
          if (hoursString.length != 2) throw 'Error @ ' + curSheet.getRange(row+1, col+2).getA1Notation() + '. The hours may be improperly formated.';
          if (isNaN(hoursString[0]) || isNaN(hoursString[1])) throw 'Error @ ' + curSheet.getRange(row+1, col+2).getA1Notation() + 
                                                                   '. The hours may contain text, or are improperly formatted.';
          if (hoursString[1] - hoursString[0] < 0 || hoursString[1] - hoursString[0] > 12 || hoursString[1] < hoursString[0]) {
            throw 'Error @ ' + curSheet.getRange(row+1, col+2).getA1Notation() + '. Reading that the employee has worked ' + (hoursString[1] - hoursString[0]) +
              ' hours on the day containing cell ' + curSheet.getRange(row+1, col+2).getA1Notation() + 
                '. Please make sure the hours are correctly formatted and in 24r time.';
          }
          // If there are no errors add the hours.
          emList[nameId].addHours(hoursString[1] - hoursString[0]);
          
          row = rowReset;
          break;
        } else {
          row++;
        }
      }
    }
    
    // Reset the row back to the top.
    row = rowReset;    
    // Change sheets at the end of the month.
    if (startCal[row-1][col] == lastDay) {
      curSheet = eSheet;
      curCal = endCal;
      row = 2;
      col = 0;
      rowReset = row;
    // At the end of a calendar column go to the next row.
    } else {
      col += 2;
      if (col > 12) {
        col = 0;
        row += 10;
        rowReset = row;
      }
    }
  }
  return emList[nameId].hours;
}

/**
* Add the pay period formula to the spreadsheet.
*/
function calcPay(startPay, endPay) {
  var ss = SpreadsheetApp.getActive();
  
  var startArray = startPay.split('-');
  var endArray = endPay.split('-');
  
  var startDate = new Date(startArray[0], startArray[1]-1, startArray[2], 0, 0, 0, 0),
      endDate = new Date(endArray[0], endArray[1]-1, endArray[2], 0, 0, 0, 0);
  
  // Use the dates to find the start and end sheet names.
  var sSheetName = getSheetDate(startDate),
      eSheetName = getSheetDate(endDate);
  
  var sStart = ss.getSheetByName(sSheetName);
  var sEnd = ss.getSheetByName(eSheetName);
  
  var startCal = getCalValues(sSheetName),
      endCal = getCalValues(eSheetName);
  
  var sCalA = 'A1:N' + startCal.length;
  var eCalA = 'A1:N' + endCal.length;
  
  
  var emList = createEmployeeList();
  
  var row = 57;
  if (sEnd.getRange(52, 1).getValue()) row = 67;
  var col = 2;
  while (sEnd.getRange(row-1, col).getValue()) col += 2;
  
  // Add the start and End of the pay period to the spreadsheet;
  sEnd.getRange(row-1, col).setValue(startPay).setNumberFormat("MMM d, yyyy");
  sEnd.getRange(row-1, col+1).setValue(endPay).setNumberFormat("MMM d, yyyy");
  
  //Get the A1 notation from start and end pay feilds.
  var sPayA = sEnd.getRange(row-1, col).getA1Notation(),
      ePayA = sEnd.getRange(row-1, col+1).getA1Notation();
  
  var formulas = new Array(emList.length);
  
  // Add formulas to the the array;
  for (var i=0, rowi = row; i<emList.length; i++, rowi++) {
    formulas[i] = ["= calcPayFormula(A" + rowi + "," + sPayA + "," + ePayA + ",'" + sSheetName + "'!" + sCalA + ",'" + eSheetName + "'!" + eCalA + ")"];
  }
  
  sEnd.getRange(row, col, emList.length, 1).setFormulas(formulas).setHorizontalAlignment('right');
}

/**
* Remove a the date and formulas from a pay Period
*
* @param {String} startDate The name of the sheet to put the template on.
*/
function deletePayPeriod(payPeriod) {
  var ss = SpreadsheetApp.getActive(),
      s = ss.getActiveSheet();
  
  // Find the starting row and find the pay period to remove.
  var row = 54;
  if (s.getRange(52, 1).getValue()) row = 64;
  var col = payPeriod*2;
  
  // Employee list needed to get the amount of employees in the pay period.
  var emList = createEmployeeList();
  
  // Clear the pay period.
  s.getRange(row+2, col, emList.length+1, 2).clearContent();
}

/**
* These functions are here to run deletePayPeriod() from the menu.
*/
function rePPOne() {
  deletePayPeriod(1);
}
function rePPTwo() {
  deletePayPeriod(2);
}
function rePPThree() {
  deletePayPeriod(3);
}

/**
* Makes the template for the possible pay periods.
*
* @param {String} startDate The name of the sheet to put the template on.
*/
function ppTemplate(startDate) {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(startDate);
  
  // Start of the employee list on the sheet.
  var stl = 56;
  if (s.getRange(52, 1).getValue()) stl = 66;
  
  // Add the title.
  s.getRange(stl, 1).setValue('Employees').setFontWeight('bold').setBackground('#336600').setFontColor('white');
  
  // Make the list of employees.
  var emList = createEmployeeList(startDate);
  
  // Add the template for the pay periods.
  for (var i=1; i<=3; i++) {
    s.getRange(stl-2, i*2, 1, 2).merge().setFontWeight('bold').setHorizontalAlignment('center').setValue('Pay Period ' + i).
    setBorder(true, true, true, true, null, null).setBackground('#336600').setFontColor('white');
    s.getRange(stl-1, i*2, 2, 2).setBorder(true, true, true, true, null, null);
    s.getRange(stl-1, i*2).setValue('Start').setFontWeight('bold');
    //s.getRange(stl-1, i*2).setFormulaR1C1("=dateValidationStart(R[" + 1 + "]C[" + 0 + "])");
    s.getRange(stl-1, i*2+1).setValue('End').setFontWeight('bold');
    //s.getRange(stl-1, i*2+1).setFormulaR1C1("=dateValidationEnd(R[" + 1 + "]C[" + 0 + "])");
    s.getRange(stl+1, i*2, emList.length, 2).mergeAcross().setBorder(true, true, true, true, null, null).setHorizontalAlignment('right');
  }
  
  // Add the employee names to the sheet.
  for (var i=0; i<emList.length; i++) {
    s.getRange(i+stl+1, 1).setValue(emList[i].name);
  }
}


/**
* Get the values in the range of the calender
*
* @param {String} startDate The name of the calender sheet to get the range of.
* @return {Array} The array of data in the calender.
*/
function getCalValues(startDate) {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(startDate);
  
  var rows = 51;
  if (s.getRange(52, 1).getValue()) rows = 61;
  
  var cal = s.getRange(1, 1, rows, 14).getValues();

  return cal;
}

/**
* Get the range of the calender.
*
* @param {String} startDate The name of the calender sheet to get the range of.
* @return {Array} The array of data in the calender.
*/
function getCalRange(startDate) {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(startDate);
  
  var rows = 51;
  if (s.getRange(52, 1).getValue()) rows = 61;
  
  var cal = s.getRange(1, 1, rows, 14);

  return cal;
}

/**
* Find a months last date.
*
* @param {String} startDate The name of the month to find the last day of.
* @return {Number} The month's last date
*/
function monthLastDay(startDate) {
  var dateArray = startDate.split(" ");
  
  var endDay = 1;
  
  //Find the months last day.
  for (var i=0; i<cal_months_labels.length; i++) {
    if (cal_months_labels[i].toLowerCase() == dateArray[0].toLowerCase()) {
      endDay = cal_days_in_month[i];
    }
  }
  
  // Check if it is a leap year.
  if (endDay == 28) {
    if ((dateArray[1] % 4 == 0 && dateArray[1] % 100 != 0) || dateArray[1] % 400 == 0){
      endDay = 29;
    }
  }
  return endDay;
}

/**
* Sets data validation for a new month so employee names match the settings page.
*
* @param {String} startDate The sheet to apply the validation to.
*/
function employeeVal(startDate) {
  var ss = SpreadsheetApp.getActive();
  var set = ss.getSheetByName('Settings');
  var s = ss.getSheetByName(startDate);
  
  // Find the first day.
  var row = 2;
  while (set.getRange(row, 1).getValue()) {
    row++;
  }
  var emValList = set.getRange(2, 1, row-2).getValues();
  
  var cal = getCalValues(startDate);
  
  // Build the validation rule that checks the employee names.
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(emValList, false).build();
  
  var lastDay = monthLastDay(startDate);
  
  var locked = false;
  var col = 1;
  row = 2
  // Loop through the calendar to set validation to the areas only names should be.
  while (!locked) {
    // Locks the loop from running when the last day is hit.
    if(cal[row-1][col-1] == lastDay) locked = true;
    
    if(cal[row-1][col-1] != 0) {
      s.getRange(row+1, col, 9).setDataValidation(rule);
      dayLock = true;
    }
    
    col += 2;
    if(col > 13) {
      col = 1;
      row += 10;
    }
  }
}

/**
* Add a new employee to the current sheet and settings page.
*/
function addEmployee() {
  var name = Browser.inputBox('Enter the name of the employee.');
  if(name == 'cancel') return;
    
  var ss = SpreadsheetApp.getActive();
  var s = ss.getActiveSheet();
  var set = ss.getSheetByName('Settings');
  
  // If the user is on the setting sheet just add to the end of the employee list.
  if (s.getSheetName() == 'Settings') {
    Browser.msgBox('Please go to the month you need the employee added.');
    return;
    
    // Gather the list of employees on the setting sheet (assumes there will never be more than 25 employees).
    var ls = set.getRange(2, 1, 25).getValues();
    // Loop through the employees.
    for (var i=0; i<ls.length; i++) {
      // At the end of the employees leave the list.
      if (ls[0][i] == null) {
        ls[0][i] = name;
        set.getRange(2, 1, 25).setValues(ls);
        return;
      }
      // Crete a message is there is already an employee with that name.
      if (ls[0][i] == name) {
        Browser.msgBox('There is already an employee with that name entered.');
        return;
      }
    }
  }
  
  // Get the current employee list.
  var emList = createEmployeeList('Settings');
  
  // Check if the employee is already in the list.
  for (var i=0; i<emList.length; i++) {
    // If they are stop serching and move on
    if (emList[i].name == name) {
      break;
    // If the employee was not found add their name;
    } else if (i == emList.length-1 && emList[i] != name){
      set.getRange(emList.length+2, 1).setValue(name);
      var employee = new Employee(ss, name, [false,false,false,false,false,false,false], 0);
      emList.push(employee);
    }
  }
  
  // Clear the formating for the pay periods and add a new employee
  var row = 54;
  if (s.getRange(52, 1).getValue()) row = 64;
  s.getRange(row, 1, row+emList.length+2, 7).clearFormat();
  s.getRange(row, 11, emList.length+2, 4).clearFormat();
  
  // Add a new template to the current sheet.
  ppTemplate(s.getSheetName());
  holPayTemp(s.getSheetName());
  employeeVal(s.getName());
  
  
}

/**
* Test if an array of work days is in the correct format.
*
* @param {String[]} workDayArray The array of works days.
* @param {String} name The name of the employee.
* @return {Boolean} True is there was an issue found.
*/
function testWorkDays(workDayArray, name) {
  if (workDayArray.length <= 1 && workDayArray[0] == '') return false;
  for (var i=0; i<workDayArray.length; i++) {
    if (workDayArray[i].length != 2) {
      Browser.msgBox('Employee ' + name + ' has a imporperly formated Work Day feild.');
      return true;
    }
  }
  return false;
}


/**
* Find the Days an employee has worked in a given time period.
*
* @param {String} name The name of the employee you want to find the days of.
* @param {String} startCheck The start of the period to check.
* @param {String} endCheck The end of period to check.
* @param {String[][]} startRange The starting calendar range.
* @param {String[][]} endRange The range of the ending calendar range.
* @return {Number} The amount of days the employee has worked.
* @customfunction
*/
function holDayChecker(name, startCheck, endCheck, startRange, endRange) {
  var ss = SpreadsheetApp.getActive()
  
  // Date Object for the start and ending dates.
  var startDate = new Date(startCheck),
      endDate = new Date(endCheck);
  
  // Use the dates to find the start and end sheet names.
  var sSheetName = getSheetDate(startDate),
      eSheetName = getSheetDate(endDate);
  
  // Get the two sheets that need to be processed.
  var sSheet = ss.getSheetByName(sSheetName),
      eSheet = ss.getSheetByName(eSheetName),
      curSheet = sSheet;

  // Get the values of the start and end calendars
  var startCal = startRange,
      endCal = endRange,
      curCal = startCal;
  
  // Find the last day of the starting sheet.
  // Only needed if the pay period is over two months.
  var lastDay = monthLastDay(sSheetName);
  
  // Total days the employee has worked.
  var totalDays = 0;
  
  var row = 2,
      col = 0,
      rowReset = row;
  var locked = false,
      inPay = false;
  while (!locked) {
    rowReset = row;
    // End when the last day has been hit.
    if (curCal[row-1][col] == endDate.getDate() && curSheet.getName() == eSheet.getName()) locked = true;
    
    if (startCal[row-1][col] == startDate.getDate()) {
      inPay = true;
    }
    
    // Only check hours while in the pay period.
    if(inPay) {
      while (row < rowReset + 8) {
        if (curCal[row][col] == name && curCal[row][col+1] != '') {
          totalDays++;
          row = rowReset;
          break;
        } else {
          row++;
        }
      }
    }
    
    // Reset the row back to the top.
    row = rowReset;    
    // Change sheets at the end of the month.
    if (startCal[row-1][col] == lastDay) {
      curSheet = eSheet;
      curCal = endCal;
      row = 2;
      col = 0;
      rowReset = row;
    // At the end of a calendar column go to the next row.
    } else {
      col += 2;
      if (col > 12) {
        col = 0;
        row += 10;
        rowReset = row;
      }
    }
  }
  
  return totalDays;
}

/**
* Add the formulas for the holiday checker to the spreadsheet.
*
* @param startCheck The day to start checking on.
* @param endCheck The last day to end checking on.
*/
function holDayCalc(startCheck, endCheck) {
  var ss = SpreadsheetApp.getActive(),
      s = ss.getActiveSheet();
  
  // Turn dates into an array.
  var startArray = startCheck.split('-');
  var endArray = endCheck.split('-');
  
  // The arrays are used to create new date objects.
  var startDate = new Date(startArray[0], startArray[1]-1, startArray[2], 0, 0, 0, 0),
      endDate = new Date(endArray[0], endArray[1]-1, endArray[2], 0, 0, 0, 0);
  
  // Use the dates to find the start and end sheet names.
  var sSheetName = getSheetDate(startDate),
      eSheetName = getSheetDate(endDate);
  
  // Get the start and end sheet to check on.
  var sStart = ss.getSheetByName(sSheetName);
  var sEnd = ss.getSheetByName(eSheetName);
  
  var startCal = getCalValues(sSheetName),
      endCal = getCalValues(eSheetName);
  
  // The range of the start and end calendar to be inputed into the formulas.
  var sCalA = 'A1:N' + startCal.length;
  var eCalA = 'A1:N' + endCal.length;
  
  
  var emList = createEmployeeList();
  
  var row = 57;
  if (s.getRange(52, 1).getValue()) row = 67;
  var col = 12;
  
  
  // Add the start and End of the pay period to the spreadsheet;
  s.getRange(row-2, col).setValue(startCheck).setNumberFormat("MMM d, yyyy");
  s.getRange(row-2, col+2).setValue(endCheck).setNumberFormat("MMM d, yyyy");
  
  //Get the A1 notation from start and end pay feilds.
  var sPayA = sEnd.getRange(row-2, col).getA1Notation(),
      ePayA = sEnd.getRange(row-2, col+2).getA1Notation();
  
  var formulas = new Array(emList.length);
  
  // Add formulas to the the array;
  for (var i=0, rowi = row; i<emList.length; i++, rowi++) {
    formulas[i] = ["=calcPayFormula(K" + rowi + "," + sPayA + "," + ePayA + ",'" + sSheetName + "'!" + sCalA + ",'" + eSheetName + "'!" + eCalA + ")"];
  }
  // Add the formulas to check total hours worked to the calendar.
  s.getRange(row, col, emList.length, 1).setFormulas(formulas).setHorizontalAlignment('right');
  
  // Overight the formulas array to store the days worked.
  for (var i=0, rowi = row; i<emList.length; i++, rowi++) {
    formulas[i] = ["=holDayChecker(K" + rowi + "," + sPayA + "," + ePayA + ",'" + sSheetName + "'!" + sCalA + ",'" + eSheetName + "'!" + eCalA + ")"];
  }
  // Add the formula for checking days worked.
  s.getRange(row, col+1, emList.length, 1).setFormulas(formulas).setHorizontalAlignment('right');
  
  // Overight the formulas array once again to store the very small average formula;
  for (var i=0, rowi = row; i<emList.length; i++, rowi++) {
    formulas[i] = ["=customDivison(R[" + 0 + "]C[" + -2 + "],R[" + 0 + "]C[" + -1 + "])"];
  }
  // Add the formula that finds average hours per day worked to the calendar.
  s.getRange(row, col+2, emList.length, 1).setFormulasR1C1(formulas).setHorizontalAlignment('right').setNumberFormat('0.00');
}

/**
* A custom divsion function because it's way better that dividing by zero equals zero!.
*
* @param {number} value1 The dividend.
* @param {number} value2 The divisor.
* @return {number} The quotient.
* @customFunction
*/
function customDivison(value1, value2) {
  return value2 == 0 ? 0 : value1/value2;
}

/**
* Sets up the section for calculated days worked for holiday pay.
*
* @param {String} startDate The name of the sheet to put the template on.
*/
function holPayTemp(startDate) {
  var ss = SpreadsheetApp.getActive(),
      s = ss.getSheetByName(startDate);
  
  var emList = createEmployeeList('Settings');
  
  var row = 54;
  if (s.getRange(52, 1).getValue()) row = 64;
  
  // Merge the required cells and put the correct borders down.
  s.getRange(row, 11, 1, 4).merge().setBorder(true, true, true, true, false, false).setHorizontalAlignment('center');
  s.getRange(row+1, 11, 1, 2).setBorder(true, true, true, true, false, false);
  s.getRange(row+1, 13, 1, 2).setBorder(true, true, true, true, false, false);
  s.getRange(row+2, 11, 1, 4).setBorder(true, true, true, true, true, false);
  s.getRange(row+3, 11, emList.length, 4).setBorder(true, true, true, true, true, false);
  
  var range = s.getRange(row, 11, 3, 4);
  // Set the values and format them in the perment cells.
  range.setValues([['Holiday Time Calculator','','',''], 
                   ['Start','','End',''], 
                   ['Employees','Total Hours','Days Worked','Average']]);
  range.setBackgrounds([['#336600','','',''],['','','',''],['#336600','','','']]);
  range.setFontColors([['white','','',''],['','','',''],['white','','','']]);
  range.setFontWeights([['bold','','',''],['bold','normal','bold','normal'],['bold','bold','bold','bold']])
  
  range = s.getRange(row+3, 11, emList.length)
  var values = range.getValues();
  
  for (var i=0; i<emList.length; i++) {
    values[i][0] = emList[i].name;
  }
  
  range.setValues(values);
  
  //range = s.getRange(row+1, 11, 1, 3);
  //range.setFormulasR1C1("=dateValidationStart(R[" + 0 + "]C[" + 1 + ")","","=dateValidationEnd(R[" + 0 + "]C[" + 1 + ")");
}

/**
* Checks the list of works hours to make sure the format is correct.
*
* @param {'10-17','10-18','10-18','10-18','10-18','10-18','10-17.5'} hours A list of works hours.
* @return {String} A message that lets the user know the hours are the correct format.
* @customFunction
*/
function storeHoursValidation(hours) {
  for (var row = 0; row < hours.length; row++) {
    for (var col = 0; col < hours[row].length; col++) {
      var hrArray = hours[row][col].split('-');
      
      // Make sure the hours are numbers.
      if (isNaN(hrArray[0]) || isNaN(hrArray[1])) throw hours[row][col] + ' contains letters or symbols. Must only be numbers seperated by a "-".';
      // Make sure the array is of length two.
      if (hrArray.length != 2) throw hours[row][col] + ' is not in the proper format. ex) 10-18';
      // Make sure there is not negative hours in a day
      if (hrArray[1] - hrArray[0] < 0) throw hours[row][col] + ' would result in a day being negative time. Make sure values are in 24hr time. ex) 10-18';
    }
  }
  return 'Good';
}

/**
* Checks the list of work days to make sure the format is correct.
*
* @param {String} days A list of work days.
* @return {String} A message that lets the user know the work days are the correct format.
* @customFunction
*/
function workDaysValidation(days) {
  for (var row = 0; row < days.length; row++) {
    for (var col = 0; col < days[row].length; col++) {
      var dayArray = days[row][col].split(',');
     
      // Check if there is more than 7 days writen up.
      if (dayArray.length > 7 ) throw 'An employee is set to work more than 7 days.';
      
      // Loop through the days for problems.
      for(var i=0; i<dayArray.length; i++) {
        // Check if there is a comma followed by nothing.
        if (dayArray[i] == '' && i != 0 && dayArray[i+1] != '') throw 'Error: ' + dayArray[i] + ' Please make sure days are in two character short form, and there is no trailing commas.';
        // Check and make sure all days are length 2.
        if (dayArray[i].length != 2 && i != 0) throw 'Error: ' + dayArray[i] + ' Please make sure days are in two character short form, and there is no trailing commas.';
      }
      
      // Make sure days are in order and in the list.
      if (dayArray[0] != '') {
        var cor = false;
        for (var i=0, j=0; j<dayArray.length && i<cal_work_days.length; i++) {
          // Check if there is a match in the list.
          if (dayArray[j] == cal_work_days[i]) {
            cor = true;
            j++;
          } else {
            // False if there is no match.
            cor = false;
          }
          // Make sure no days follow saturday.
          if (dayArray[j] == 'sa' && j != dayArray.length - 1) {
            cor = false;
            break;
          }
        }
        // Give and error if there is no matches to a day.
        if (!cor) throw 'Days are in incorrect order or do not match the list.';
      }
    }
  }
  return 'Good';
}

/**
* Checks dates to make sure they are formatted correctly
*
* @param {String} input A date to be checked.
* @return {String} Returns Start if the date is correct
* @customFunction
*/
function dateValidationStart(input) {
  // Say 'Start' if there is no value.
  if (input == '') return 'Start';
  
  try {
    var dayArray = [input.getMonth()+1, input.getDate(), input.getFullYear()];
  } catch(err) {
    throw 'Please make sure the date is formatted mm/dd/yyyy.';
  }  
  
//  if (dayArray.length != 3) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  if (isNaN(dayArray[0]) || isNaN(dayArray[0]) || isNaN(dayArray[0])) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  if (dayArray[0] < 1 || dayArray[0] > 12) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  if (dayArray[1] < 1 || dayArray[1] > 31) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  // Sorry, the calender won't work after the year 3000.
//  if (dayArray[2] < 2015 || dayArray[2] > 3000) throw 'Please make sure the date is formatted mm/dd/yyyy.';
  
  return 'Start';
}

/**
* Checks dates to make sure they are formatted correctly
*
* @param {String} input A date to be checked.
* @return {String} Returns Start if the date is correct
* @customFunction
*/
function dateValidationEnd(input) {
  // Say 'End' if there is no value.
  if (input == '') return 'End';
  
  try {
    var dayArray = [input.getMonth()+1, input.getDate(), input.getFullYear()];
  } catch(err) {
    throw 'Please make sure the date is formatted mm/dd/yyyy.';
  }
  
//  if (dayArray.length != 3) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  if (isNaN(dayArray[0]) || isNaN(dayArray[0]) || isNaN(dayArray[0])) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  if (dayArray[0] < 1 || dayArray[0] > 12) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  if (dayArray[1] < 1 || dayArray[1] > 31) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  // Sorry, the calender won't work after the year 3000.
//  if (dayArray[2] < 2015 || dayArray[2] > 3000) throw 'Please make sure the date is formatted mm/dd/yyyy.';
//  
  return 'End';
}

/**
* Check if a day is a holiday given the month and year. 
*
* @param {String} startDate Name of the sheet but also contains year and month.
* @return {Number} Returns the day the holiday is on.
*/
//function testForHoliday(startDate) {
//  var dayArray = startDate.split(' ');
//  var date = new Date(dayArray[1], 0, 1);
//
//  switch (dayArray[0]) {
//    case 'January':
//      return 1;
//    case 'February':
//      date.setMonth(1);
//      return 1+8-date.getDay();
//    case 'March':
//      date.setMonth(2);
//      return
//  }
//  
//}

/**
* A test function to help test new methods for debugging.
*/
function test() {
  employeeVal('January 2016');
  //Logger.log(SpreadsheetApp.getActiveSheet().getRange(2, 1, 25).getValues());
  //Logger.log(testForHoliday('February 2015'));
  //var startDate = 'February 2016';
  //addHours(startDate);
  //deletePayPeriod(2)
  //var workDaysSet = SpreadsheetApp.getActive().getSheetByName('Settings').getRange(row, 2).getValue().split(',');
  //createCal('May 2016');
  //for(var i=1; i<=14; i++) SpreadsheetApp.getActiveSheet().setColumnWidth(i, 85);
  //holPayTemp('January 2016');
  //holDayCalc('2016-5-1', '2016-5-30')
  //Logger.log(SpreadsheetApp.getActiveSheet().getRange('L65').getValue())
}
