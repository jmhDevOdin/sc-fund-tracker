/*
 * Help From:
 * https://gist.github.com/paulgambill/cacd19da95a1421d3164
 * http://docs.brightcove.com/en/video-cloud/analytics-api/samples/spreadsheet-script.html
 *
 * https://mashe.hawksey.info/2012/09/google-app-script-scheduling-timed-triggers/
 * 
 * Custom made for purpose by Jonathan Harrison - @j-m-harrison on gitlabModified
 * Implement: https://developers.google.com/apps-script/guides/triggers/    Probably onGet
 */

//Global configuration variables
var URL = 'http://pledgetrack.rabbitsraiders.net?json=1'; //datasource
var rsiUrl = 'https://robertsspaceindustries.com/api/stats/getCrowdfundStats';
var FIRST_ROW = 19; //location of first data row
var FIRST_DATE = 1367366400; //epoch time for 2013-05-01
var FIRST_FUNDING_COLUMN = 'B'; //Name of first data column
var FIRST_FUNDING_COLUMN_INDEX = 2; //Index of first data column
var NUM_HOURLY_COLS = 25; //Number of data columns
var FUNDING_TO_CITIZEN_OFFSET = 26;
var FUNDING_TO_FLEET_OFFSET = 52;
var SECOND_IN_DAY = 86400000; //duh
var UTC1_TO_GMT_OFFSET = -1; //time difference between UTC+1 and GMT

var hourlySheet; //used for caching teh hourly sheet object
var cronSheet; //used for caching teh cron sheet object
var firstDate; //used for caching teh first Date

function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('RSI Pledge Data')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
  //showSidebar();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('PledgeDataControlPanel')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('RSI Pledge Data Control Panel')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

/* Retrieves JSON data from a remote webservice and updates the sheet */
function fetchCSVData() {
  var nextDateTime;
  var countPulled = 0;
  //try to get the next empty date
  try { //if the date column has run out this can throw an exception
    nextDateTime = getFirstEmptyDate();
  } catch(e) { //so add more rows to the date column and try again
    updateDateColumn();
    nextDateTime = getFirstEmptyDate();
  }
  Logger.log("Next DateTime " + nextDateTime);
  var nextDateObj = getDateFromDateString(nextDateTime); //convert to datetime obj
  //fetch data and parse it into a usable format
  var dataURL = getUrlForTimeStamp(nextDateObj);
  Logger.log("Fetch: " + dataURL);
  var response = UrlFetchApp.fetch(dataURL);
  var json = response.getContentText();
  var data = JSON.parse(json);
  var sheet = getHourlySheet();
  var writeData = false;
  if(sheet != null) {
    for(var i = 0; i < data.length; i++) {
      var row = data[i];
      var timeStamp = row.TimeStamp;
      var tsObj = getDateFromDateString(timeStamp, 1);
      if(tsObj.getTime() >= nextDateObj.getTime() && !writeData) {
        Logger.log("TimeStamp " + timeStamp);
        writeData = true;
      }
      if(writeData) {
        countPulled++;
        writeTimeEntry(tsObj, row);
      }
    }
  }
  logHourlyPull("CSVPull: Success! Pulled " + countPulled + " rows! URL: " + dataURL);
}

function getUrlForTimeStamp(dateTime) {
  var dataURL = URL;
  dataURL += '&startingDateTime='+dateTime.toISOString()+'&offset=-1';
  
  return dataURL;
}

/**
 * Writes the data table for the specified row using the specified values
 *
 * @param Date timeObj  Date to update data for
 * @param Array row  Array of data, must include 'funding', 'citizens', and 'fleet'
 */
function writeTimeEntry(timeObj, row) {
  var cellRange = getCellsByTimestamp(timeObj);
  //Update funding
  if('funding' in cellRange && 'Funding' in row) {
    var fundingRanges = cellRange.funding;
    for(var i = 0; i < fundingRanges.length; i++) {
      var range = fundingRanges[i];
      range.setValues([ [ row.Funding ] ]);
    }
  }
  //Update citizens
  if('citizens' in cellRange && 'Citizens' in row) {
    var citizenRanges = cellRange.citizens;
    for(var i = 0; i < citizenRanges.length; i++) {
      var range = citizenRanges[i];
      range.setValues([ [ row.Citizens ] ]);
    }
  }  
  //Update fleet
  if('fleet' in cellRange && 'Fleet' in row) {
    var fleetRanges = cellRange.fleet;
    for(var i = 0; i < fleetRanges.length; i++) {
      var range = fleetRanges[i];
      range.setValues([ [ row.Fleet ] ]);
    }
  }
    
}

/**
 * Convert a Datestring into a Date Object
 *
 * @param String  Date String like  yyyy-mm-dd hh:mm:ss
 * @param Integer offset  An integer representing a timezone offset in hours
 *
 * @return Date a Date object for this date string
 */
function getDateFromDateString(dateStr, offset) {
  offset = offset | 0;
  var parts = dateStr.split(' ');
  var date = parts[0];
  var time = parts[1];
  var dateParts = date.split('-');
  var timeParts = time.split(':');
  
  var dateObj = new Date(Date.UTC(dateParts[0], dateParts[1] - 1, dateParts[2], timeParts[0], timeParts[1], timeParts[2]));
  if(offset > 0) {
    dateObj.setTime(dateObj.getTime() + (offset * 3600 * 1000));
  }
  return dateObj;
}

/**
 * Finds the Sheet Object for our hourly data capture
 *
 * @return Sheet  The hourly data capture sheet
 */
function getHourlySheet() {
  if(hourlySheet == null) { //if we don't have this precached
    //get all sheets
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for(var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      //look for the sheet with the matching name
      if(sheet.getName() == 'Hourly Pledge Capture'){
        hourlySheet = sheet; //cache value
      }
    }
  }
  
  return hourlySheet;
}

/**
 * Finds the Sheet Object for our hourly data capture
 *
 * @return Sheet  The hourly data capture sheet
 */
function getCronLogSheet() {
  if(cronSheet == null) { //if we don't have this precached
    //get all sheets
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for(var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      //look for the sheet with the matching name
      if(sheet.getName() == 'Cron Log'){
        cronSheet = sheet; //cache value
      }
    }
  }
  
  return cronSheet;
}

/** 
 * Finds the first Timestamp in the table
 * 
 * @return Date  firstDate 
 */
function getFirstDate() {
  if(firstDate == null) { //if we don't have this precached
    var derpOffset = new Date().getTimezoneOffset() * 60; // because javascript auto timezones shit * 60 because timezone offset is in minutes for some reason
    var firDate = (FIRST_DATE + derpOffset) * 1000; // * 1000 because javascript epochs are in milliseconds... for some reason
    //get the value in the date column in the first row
    //var sheet = getHourlySheet();
    //var hourValues = sheet.getRange(FIRST_ROW, FIRST_FUNDING_COLUMN_INDEX - 1, 1, 1).getValues();
    firstDate = firDate;//hourValues[0][0]; //cache value
  }
    
  return firstDate;
}

/**
 * Gets the a CellRange object representing the cells
 * for this timestamp
 *
 * @param Date second  The timestamp to get cells for
 *
 * @return Array<CellRange>  An array of arrays of cellranges, indexed by datatype (funding, citizens, fleet)
 *                           First array indexes datatype, second array is a collection of cells for that datatype
 */
function getCellsByTimestamp(second) {
  var first = getFirstDate();
  var dayDiff = Math.floor((second-first)/(1000*60*60*24)); //get the number of millisceconds between, divide to days
  var hours = (second.getHours() + (second.getTimezoneOffset() / 60)); // get hours form the date, add the timezone offset and module by number of hours in a day
Logger.log("day diff " + dayDiff + "  hours " + hours);
  var dayOffset = Math.floor(hours / 24);
  hours = hours % 24;
  var row = FIRST_ROW + dayDiff + dayOffset;
  Logger.log("hours after fix " + hours + " row offset " + dayOffset);
  var col = getColumnIndexFromHour(hours);
  Logger.log("col from hours" + hours + " " + col);
  Logger.log("Cell at " + row + ":" + col);

  var sheet = getHourlySheet();
  var dateRange = sheet.getRange(row, FIRST_FUNDING_COLUMN_INDEX - 1, 1, 1);
  var dateValue = dateRange.getValues()[0][0];
  if(dateValue == null || dateValue == '') {
    updateDateColumn();
  }
                                   
  var fundingRange = [ sheet.getRange(row, col, 1, 1) ];
  var citizenRange = [ sheet.getRange(row, col + FUNDING_TO_CITIZEN_OFFSET, 1, 1) ];
  var fleetRange   = [ sheet.getRange(row, col + FUNDING_TO_FLEET_OFFSET, 1, 1) ];
  //if this is 00:00:00, update the 24:00:00 column on the day before
  if(col == FIRST_FUNDING_COLUMN_INDEX) {
    fundingRange.push(sheet.getRange(row - 1, NUM_HOURLY_COLS + 1, 1, 1));
    citizenRange.push(sheet.getRange(row - 1, NUM_HOURLY_COLS + 1 + FUNDING_TO_CITIZEN_OFFSET, 1, 1));
    fleetRange.push(sheet.getRange(row - 1, NUM_HOURLY_COLS + 1 + FUNDING_TO_FLEET_OFFSET, 1, 1));
  }
  
  var cellRange = { 'funding': fundingRange, 'citizens': citizenRange, 'fleet': fleetRange };
  return cellRange;
}

/**
 * Convert a cell location into a timestamp
 *
 * @BUG will fail for the first cell because of the -1 offset... oh well
 */
function getTimestampByCell(row, column, offset) {
  offset = offset || 0;
  
  var dateRow = row;

  
  var timeVals = getTimeFromColumn(column, offset);
  var hour = timeVals['hour'];
  
  if(hour < 0) {
    //figure out how many days behind we are... i'm not really sure how this works either...
    // we want 0 == 0, -1 - -24 = 1, -25 - -48 = 2 etc etc and it does!
    var dayOffset = Math.floor((Math.abs(hour) - 1) / 24) + 1;
    // get the 0-23 value of the hour
    var hour = 24 - (Math.abs(hour) % 24);
    if(hour == 24) { //24 == 0
      hour = 0;
    }
    //it's too hard to deal with day overflows so we'll just get the day from the next row up :D
    dateRow -= dayOffset;
                      
  } else if(hour > 23) {
    var dayOffset = Math.floor(hour/24); //how many days in the future is this
    hour = hour % 24; //get what hour the next day it is
    //it's too hard to deal with day overflows so we'll just get the day from the corresponding row
    dateRow += dayOffset;
  }

  var dateVals = getDateFromRow(dateRow);
  var year = dateVals['year'];
  var mon = dateVals['month'];
  var day = dateVals['day'];

  var date = (("0000" + year).slice(-4)) + "-" + (("00" + mon).slice(-2)) + "-" + (("00" + day).slice(-2));
  var time = (("00" + hour).slice(-2)) + ":00:00";
  
  var timestamp = date + " " + time;
  
  return timestamp;
}

function debugGetTimestampByCell() {
  var tests = [
    [ 905, 33 ]//, [905, 26], [905, 27], [905, 28], [905, 55]
  ];
  for(var i = 0; i < tests.length; i++) {
    var test = tests[i];
    var row = test[0];
    var col = test[1];
    var time = getTimestampByCell(row, col, UTC1_TO_GMT_OFFSET);
    Logger.log("Debug: "+row+": "+col+" : " + time);
  }
}

function test_getDateFromRow() {
  var row = 2197;
  getDateFromRow(row);
}

function getDateFromRow(row) {
  if(row < FIRST_ROW) {
    throw "Row " + row + " is not a valid data-entry row";
  }
  //var dateCol = FIRST_FUNDING_COLUMN_INDEX - 1; //Date is the 1st column  
  //var sheet = getHourlySheet();
  //var dateRange = sheet.getRange(row, dateCol, 1, 1);
  //var dateValue = dateRange.getValues()[0][0];
  var dayDiff = row - FIRST_ROW; //Number of days between first day and this row
  //Here, we get the Hardcoded Epoch time for the first date in the sheet's first row
  // Then we add our timezone offset, because otherwise we get dicked over by locality (YAY)
  var first = getFirstDate();
  var thisDate = new Date(first);
  var tdd = thisDate.getDate();
  // NOw we have a date object that is equal to our first orw
  // add a number of days equal to the difference in rows
  thisDate.setDate(thisDate.getDate() + dayDiff);
  var year = thisDate.getFullYear();
  var mon = thisDate.getMonth() + 1;
  var day = thisDate.getDate();
  
  return { year: year, month: mon, day: day };
}

function getTimeFromColumn(column, offset) {
  offset = offset || 0;
  if(column < FIRST_FUNDING_COLUMN_INDEX || column > FUNDING_TO_FLEET_OFFSET + NUM_HOURLY_COLS) {
    throw "Column " + column + " is not a valid data-entry column";
  }
  //var sheet = getHourlySheet();  
  //var timeRow = FIRST_ROW - 1; // Time is the 1st row  
  //var timeRange = sheet.getRange(timeRow, column, 1, 1);
  //var timeValue = timeRange.getValues()[0][0];
  //var hour = timeValue.getHours() + offset;
  var hour = column - FIRST_FUNDING_COLUMN_INDEX; //FIRST FUNDING Column == 00
  while(hour > 24) {
    hour -= 26;
  }
  hour += offset;
  
  return { hour: hour };
}

/**
 * Finds the Column Index for the given hour
 *
 * @param hour The hour to convert
 *
 * @return int columnIndex  The column index
 */
function getColumnIndexFromHour(hour) {
  return hour + 2;
}

/**
 * Finds the hour from the given ColumnIndex
 *
 * @param int columnIndex  The columnIndex to convert
 *
 * @return int hour  The hour
 */
function getHourFromColumnIndex(columnIndex) {
  if(columnIndex > 1) {
    return columnIndex - 2;
  } else {
    return null;
  }
}

/**
 * Converts a Date object into a Date string
 *
 * @param Date dateObj  The date to convert
 *
 * @return string  The date string for this Date
 */
function getDateString(dateObj) {
  var month = dateObj.getUTCMonth() + 1; //months from 1-12
  if(month < 10) { //0 pad month if 1 digit
    month = "0" + month;
  }
  var day = dateObj.getUTCDate();
  if(day < 10) { //0 pad day if 1 digit
    day = "0" + day;
  }
  var year = dateObj.getUTCFullYear();
  var str = year + '-' + month + '-' + day;
  
  return str;
}

/**
 * Finds the first empty row in the specified column
 *
 * @param int column  The index of the column to search
 * @param Sheet sheet  The Sheet to search
 *
 * @return int  The row number of the first empty column
 */
function getFirstEmptyRowInColumn(column, sheet) {
  var rowThrottle = 100; //Search 100 rows at a time
  var start = FIRST_ROW; //Starting index should be our first row
  // Basic algorithm, check the last element of each step.
  // If it's not empty, it's safe to assume the whole range is full so increment
  do {
    start += rowThrottle; 
    var range = sheet.getRange(start, column);
    var value = range.getValues();
    Logger.log("Trying " + start + " x " + column + ": " + value);
  } while(value != "");
  //Once we find an empty cell, we get the previous 100 cells (size of our search space) and look for a blank one
  var range = sheet.getRange(start - rowThrottle, column, rowThrottle + 1);
  var values = range.getValues();
  var ct = 0;
  while ( values[ct][0] != "" ) {
  Logger.log(ct);
    ct++;
  }
  
  //This is the first row with a blank cell
  var row = start - rowThrottle + ct;
  
  return row;
}

/**
 * Searchs the hourly sheet for the first row that hasn't been prefilled
 *
 * @return string Datestring of the first empty cell
 */
function getFirstEmptyDate() {
  //Basic algorithm: we check every 100 rows for blank midnight cells
  //This helps us find the empty cell first
  var sheet = getHourlySheet();
  var date;
  var row = getFirstEmptyRowInColumn(FIRST_FUNDING_COLUMN_INDEX, sheet);
  if(row > FIRST_ROW) { //if we have no data, this will be <= FIRST_ROW
    //Now we've found the first blank row, we search the row 1 above us because it might have empty values
    var dateRow = sheet.getRange(row - 1, FIRST_FUNDING_COLUMN_INDEX - 1, 2, NUM_HOURLY_COLS + 1); // -1 and +1 cause we want the date column too
    var hourValues = dateRow.getValues();
    var hourPtr = 1; //skip first column, it's a date
    while(hourValues[0][hourPtr] != "" && hourPtr < hourValues[0].length) {
      hourPtr++;
    }
    // if the ptr is the same as the number of columns, means the row is filled or missing the 24 column
    // which is the next day
    if(hourPtr >= NUM_HOURLY_COLS) { //missing date is next day!
      Logger.log('hourptr overflow');
      date = getTimestampByCell(row, FIRST_FUNDING_COLUMN_INDEX);
    } else {
      hourPtr++; //The values array is 0 indexed, cells are not HERP AND DERP GOOGLE
      date = getTimestampByCell(row - 1, hourPtr);
    }
    
  } else {
    //If we don't have any data, get the date in the first row
    date = getTimestampByCell(FIRST_ROW, FIRST_FUNDING_COLUMN_INDEX);
  }
  
  Logger.log("gFED: "+date);
  return date;
}

/**
 * Fills in more entires for the date column
 */
function updateDateColumn() {
  var sheet = getHourlySheet();
  var emptyDateRow = getFirstEmptyRowInColumn(FIRST_FUNDING_COLUMN_INDEX - 1, sheet);
  var hourValues = sheet.getRange(emptyDateRow - 1, FIRST_FUNDING_COLUMN_INDEX - 1, 1, 1).getValues();
  var dateObj = hourValues[0][0];
  var dateTS = dateObj.getTime();
  for(var days = 1; days < 8; days++) {
    var newTS = dateTS + (days * SECOND_IN_DAY);
    var newDate = new Date(newTS);
    //newDate.setTime( dateTS + (days * SECOND_IN_DAY) );
    sheet.getRange(emptyDateRow - 1 + days, 1, 1, 1).setValues([
        [ getDateString(newDate) ]
      ]);
  }
}

/**
 * Fetchs the Current RSI Stats!
 */
function fetchPledgeData() {
  var headers = {
    'X-Requested-With': 'XMLHttpRequest'
  };
  var payload = {
    "fleet": true,
    "fans": true,
    "funds": true
  };
  var params = {
    'method' : 'post',
    'headers': headers,
    'payload': payload
  };
  
  var response = UrlFetchApp.fetch(rsiUrl, params);
  var json = response.getContentText();
  var data = JSON.parse(json);
  var payload;
  if(data['success'] == 1) {
    payload = data.data;
  } else {
    throw "Failed to pull pledge stats";
  }
  
  if('funds' in payload) {
    payload.funds = Math.floor(payload.funds / 100); //is returning cents atm too
  }
  
  return payload;
}

function onHourlyUpdate() {
  Logger.log('Running Hourly Pull');
  var timeObj = new Date();
  var minutes = timeObj.getMinutes();
  minutes = 2;
  if(minutes > 5 && minutes < 55) { //only log data within 5 minutes of the hour
    Logger.log(" -- Cancelling, not close enough to the hour");
    logHourlyPull("Didn't pull because not within 5 minutes of the hour");
    return;
  }
  
  Logger.log('Pulling from RSI Site');
  var payload = fetchPledgeData();
  var dataRow = {
    'Funding': payload['funds'],
    'Citizens': payload['fans'],
    'Fleet': payload['fleet']
  };
  
  writeTimeEntry(timeObj, dataRow);
  logHourlyPull("RSIDirectPull: Pulled Successfully");
  Logger.log('Pull Complete');
}

var CRON_LOG_DATE_COL = 1;
var CRON_LOG_MESSAGE_COL = 2;
function logHourlyPull(message) {
  var datetime = new Date();
  var cronSheet = getCronLogSheet();
  var row = getFirstEmptyRowInColumn(CRON_LOG_DATE_COL, cronSheet);
  var cronRange = cronSheet.getRange(row, CRON_LOG_DATE_COL, 1, 2);
  cronRange.setValues([ [ datetime.toUTCString(), message ] ]);
}

function fixEmptyCell() {
  var message = "Nothing Happened";

  try {  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hSheet = getHourlySheet();
    if(ss.getSheetId() != hSheet.getSheetId()) {
      message = "Fix Empty Cell only works on the Hourly Capture Sheet!";
    } else {

      var failedMessage = '';
      var pulled = 0;
      var activeRange = ss.getActiveRange();
      for(var row = 1; row <= activeRange.getNumRows(); row++) {
        for(var col = 1; col <= activeRange.getNumColumns(); col++) {
          Logger.log('get ' + row + ' :: ' + col);
          var cell = activeRange.getCell(row, col);
          Logger.log("got");
      
          if(cell) {

            var results = pullEmptyCell(cell);
            if(results['success']) {
              pulled += results['count'];
            } else {
              failedMessage += results['message'] + "\n";
            }

        
          } else {
            failedMessage += "Unable to find Cell or no Cell selected!\n";
          }
      
        }
      }
      message = '';
      if(pulled > 0) {
        message = "Pulled " + pulled + " rows\n";
      }
      message += failedMessage;
      
    }
  } catch(exception) {
    message = exception.message;
  }
  
  Logger.log(message);

  return message;
}

function pullEmptyCell(cell) {
  var row = cell.getRow();
  var col = cell.getColumn();
  var timestamp = getTimestampByCell(row, col, UTC1_TO_GMT_OFFSET);
  var message = "Entry For " + timestamp;
  var success = false;
  var countPulled = 0;
        
  var dataURL = URL + '&dateTime=' + timestamp;
  var response = UrlFetchApp.fetch(dataURL);
  var json = response.getContentText();
  var data = JSON.parse(json);
  if(data.length > 0) {
    message = "got " + data.length + " rows for " + timestamp;
    for(var i = 0; i < data.length; i++) {
      var row = data[i];
      var timeStamp = row.TimeStamp;
      var tsObj = getDateFromDateString(timeStamp, 1);
      countPulled++;
      writeTimeEntry(tsObj, row);
      success = true;
    }
    message = "Pulled " + countPulled + " rows";
  } else {
    message = "No Data for " + timestamp + " GMT!";
  }
        
  return { 'count': countPulled, 'message': message, 'success': success };
}
function debugPullEmptyCell() {
  var hourlySheet = getHourlySheet();
    var tests = [
    [ 905, 33 ]//, [905, 26], [905, 27], [905, 28], [905, 55]
  ];
  for(var i = 0; i < tests.length; i++) {
    var test = tests[i];
    var row = test[0];
    var col = test[1];
    var range = hourlySheet.getRange(row, col, 1, 1);
    var cell = range.getCell(1, 1);
    var ret = pullEmptyCell(cell);
    
    return ret;
  }
}