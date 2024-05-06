/**
 * Base function to get previous status
 */
function dup1() {
  var activeSheet = SpreadsheetApp.getActiveSheet();
  var allsheets = SpreadsheetApp.getActive().getSheets()
  
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell()
  var row = cell.getRow();
  var address = SpreadsheetApp.getActiveSheet().getRange(row, 1).getValue();
  var oldStatus = [];

  if (!row) {
    return utils.constants.emptyVal;
  }

  for(var i in allsheets){
    var sheet = allsheets[i];

    
    if (utils.isSame(activeSheet.getName(), sheet.getName())) {
      // skip active sheet where added method
      continue;
    }

    var count = 1;
    var emptyLinesCount = 0;
    var isFound = false

    do {
      var addressVal = sheet.getRange(utils.constants.columns.address + count).getValue();
      var statusVal = sheet.getRange(utils.constants.columns.status + count).getValue();
      var dateVal = sheet.getRange(utils.constants.columns.date + count).getValue();

      isFound = utils.isSame(address, addressVal);
      count++;

      if (isFound && !utils.isEmpty(statusVal)) {
        oldStatus.push(utils.printStatus(sheet.getName(), dateVal, statusVal));
      }

      // calculate empty lines
      if (utils.isEmpty(statusVal)) {
        emptyLinesCount++;
      }
    } while (emptyLinesCount < utils.constants.limits.emptyLines && !isFound);

  }

  if (oldStatus && oldStatus.length > 0) {
    return oldStatus.toString();
  } else {
    return utils.constants.emptyVal;
  }
}

var utils = {
  /**
   * Constants
   */
  constants: {
    emptyVal: "-",
    columns: {
      address:  "A",
      date:     "I",
      status:   "J"
    },
    limits: {
      emptyLines: 3
    },
    datePattern: "dd.MM",
    date: {
      pattern: "dd.MM",
      timeZone: "GMT+2"
    }
  },

  /**
   * Check for empty
   */
  isEmpty: function(val) {
    return val === "" || val === undefined || val === null;
  },

/**
 * Check if values are same
 */
  isSame: function(val1, val2) {
    return val1 === val2;
  },

/**
 * Method to print status
 */
  printStatus: function(sheetName, date, status) {
    var str = "";

    if (sheetName) {
      // str += sheetName + " ";
    }

    if (date && date instanceof Date) {
      // formatt date into 12.12
      str += Utilities.formatDate(date, this.constants.date.timeZone, this.constants.date.pattern) + " ";
    }

    if (status) {
      str += status;
    }

    return str
  }
}
