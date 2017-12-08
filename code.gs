function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  /*****************************************************************************************************************
  *
  * 
  *
  *****************************************************************************************************************/
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("3CC P&L by Trader Report")
      .addItem("Run Manual Report", "menuItem1")
      //.addSeparator()
      //.addItem("Enable Weekly Reporting", "menuItem2")
      //.addItem("Disable Weekly Reporting", "menuItem3") 
      //.addItem("Test Auto Trigger", "triggerGenerateReports")
      .addToUi();
}

function menuItem1() {
  Logger.log("**** Running menuItem1() ****");
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt('Create Manual Report', 'Enter month and year (e.g. August 2017)', ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    var monthAndYear = response.getResponseText();
    var isManualReport = true;
    if (monthAndYear == "") {
      throw 'Error:  Must enter a month and year to run report';
    }
    dataFromGenerateReports = generateReports(monthAndYear, isManualReport);
    if (dataFromGenerateReports.dataForEmails == null) {
      ui.alert('Manual report completed successfully, but there are no loads ready for reporting in ' + monthAndYear);
    } else {
      ui.alert('Manual report for ' + monthAndYear + ' is complete.');
    }  
  }
}

function menuItem2() {
  Logger.log("**** Running menuItem2() ****");
  // Trigger every Monday at 01:00AM CT.
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert('Please confirm',
                        'Are you sure you want to enable automatic weekly reporting?',
                         ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    // Check if any time triggers have already been created
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getTriggerSource() == ScriptApp.TriggerSource.CLOCK) {
        throw 'Error:  Weekly reporting is already enabled.  Can only enable one time.';
      }
    }
    ScriptApp.newTrigger('triggerGenerateReports')
    .timeBased()
    .everyMinutes(1)
    //.onWeekDay(ScriptApp.WeekDay.MONDAY)
    //.atHour(1)
    //.inTimezone("America/Chicago")
    .create();
  
    ui.alert('Weekly reporting has been enabled.  Reports are updated every Monday at 1:00 am CT.');
  } else {
    // User clicked "No" or X in the title bar.
  }
}

function menuItem3() {
  Logger.log("**** Running menuItem3() ****");
  var ui = SpreadsheetApp.getUi(); // Same variations.
  var result = ui.alert('Please confirm',
                        'Are you sure you want to disable automatic weekly reporting?',
                         ui.ButtonSet.YES_NO);
  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    // Loop over all triggers.
    var allTriggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < allTriggers.length; i++) {
      // If the current trigger is a time trigger, delete it.
      if (allTriggers[i].getTriggerSource() == ScriptApp.TriggerSource.CLOCK) {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    }
    ui.alert('Automatic weekly reporting has been disabled.  Weekly reports will not be updated.');
  } else {
    // User clicked "No" or X in the title bar.
  }
}

function getSettings() {
  Logger.log("**** Running getSettings() ****");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Settings");
  
  if (sheet == null) {        // if Settings sheet does not exist in destination
      throw 'Error:  The Settings sheet could not be found';
  } 
  var data = sheet.getDataRange().getValues();
  
  var ssUrl = ss.getUrl();
  
  var emailsWeeklyReport = data[4][4];
  var totalNumTraders = data[5][4];
  
  var masterHeaderRowNum = data[26][4];
  var masterDataStartRowNum = data[27][4];
  var masterDataStartColNum = letterToColumn(data[28][4]);
  var masterDataEndColNum = letterToColumn(data[29][4]);
  
  var arrayTraderNames = data[53][4].split(',');
  var arrayMasterUrls = data[54][4].split(',');
  
  return {ssUrl:ssUrl,
          
          emailsWeeklyReport:emailsWeeklyReport,
          totalNumTraders:totalNumTraders,
          
          masterHeaderRowNum:masterHeaderRowNum,
          masterDataStartRowNum:masterDataStartRowNum,
          masterDataStartColNum:masterDataStartColNum,
          masterDataEndColNum:masterDataEndColNum,
          
          arrayTraderNames:arrayTraderNames,
          arrayMasterUrls:arrayMasterUrls};
}

function triggerGenerateReports() {
  var d = new Date();
  var monthAndYear = getMonthAndYear(d);
  var isManualReport = false;
  var dataFromGenerateReports = generateReports(monthAndYear, isManualReport);
  
  var month = d.getMonth() + 1; //months from 1-12
  var day = d.getDate();
  var year = d.getFullYear();
  var formatedDate = month + "/" + day + "/" + year;
   
  var mailTo = dataFromGenerateReports.settings.emailsWeeklyReport;
  var mailSubject = 'Weekly P&L Report - ' + formatedDate;
  if (dataFromGenerateReports.dataForEmails == null) {
    var mailBody = '<p>As of ' + formatedDate + ', there are no loads ready for reporting yet in ' + monthAndYear + '.</p>' + 
                   '<p>To view historical data, check out the <a href="' + dataFromGenerateReports.settings.ssUrl + '">3CC Profit by Trader Report</a>.</p>';
  } else {
    var mailBody = '<p>The weekly summary and detailed reports for ' + monthAndYear + ' with loads up to ' + formatedDate + ' are now available to view in the <a href="' + dataFromGenerateReports.settings.ssUrl + '">3CC Profit by Trader Report</a>.</p>' +
                   '<p>To date this month, Total Profit is $' + dataFromGenerateReports.dataForEmails.totalProfitForAllTraders + ' and Total 3CC Final Invoice Amount is $' + dataFromGenerateReports.dataForEmails.totalCcFinalInvoiceAmountforAllTraders + ' for a Profit Margin of ' + (dataFromGenerateReports.dataForEmails.totalProfitForAllTraders * 100 / dataFromGenerateReports.dataForEmails.totalCcFinalInvoiceAmountforAllTraders).toFixed(1) +  '%.</p>';
  }
  Logger.log("mailTo = " + mailTo);
  Logger.log("mailSubject = " + mailSubject);
  Logger.log("mailBody = " + mailBody);
  
  MailApp.sendEmail({
    to: mailTo,
    subject: mailSubject,
    htmlBody: mailBody,
    noReply: true
  });
}

function generateReports(monthAndYear, isManualReport) {
  Logger.log("**** Running generateReports() ****");
  
  var myActiveSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var d = new Date();
  
  myActiveSpreadsheet.toast('Loading settings', 'Status');
  var settings = getSettings();
  
  myActiveSpreadsheet.toast('Loading data from Masters', 'Status');
  var filteredObjectRowsData = getAllDataFromMasters(settings, monthAndYear);
  
  myActiveSpreadsheet.toast('Creating reports', 'Status');
  var dataForEmails = createDetailandSummaryReport(filteredObjectRowsData, settings, d, monthAndYear, isManualReport);

  return {settings:settings,
          dataForEmails:dataForEmails};
}

function getAllDataFromMasters(settings, monthAndYear) {
  Logger.log("**** Running getAllDataFromMasters ****");
  var arrayMasterUrls = settings.arrayMasterUrls;
  var arrayTraderNames = settings.arrayTraderNames;
  var masterHeaderRowNum = settings.masterHeaderRowNum;  
  
  var masterDataStartRowNum = settings.masterDataStartRowNum;
  var masterDataStartColNum = settings.masterDataStartColNum;
  var masterDataTotalCols = settings.masterDataEndColNum - masterDataStartColNum + 1;
  
  var objectRowsData = {};
  var filteredObjectRowsData = {};
  var allDataFromMasters = {};
  
  // Loop through the master Urls
  for (var i = 0; i < arrayMasterUrls.length; i++) {
    // Open the appropriate sheet
    Logger.log("monthAndYear = " + monthAndYear);
    Logger.log("arrayMasterUrls[i] = " + arrayMasterUrls[i]);
    var sheet = SpreadsheetApp.openByUrl(arrayMasterUrls[i]).getSheetByName(monthAndYear);
    if (sheet == null) {
      filteredObjectRowsData[arrayTraderNames[i]] = [];
      continue;
      //throw "Error: The " + monthAndYear + " sheet could not be found in " + arrayTraderNames[i] + "\'s Master";
    }
    // Get range of data based on user supplied settings for start row, start col, and end col
    // >>>> Eventually will need to look into whether using LastRow is an issue here, vs stripping out blank cells at bottom
    var range = sheet.getRange(masterDataStartRowNum, masterDataStartColNum, sheet.getLastRow(), masterDataTotalCols);
    
    // store row data indexed by column header name, with trader name as the key
    objectRowsData[arrayTraderNames[i]] = getRowsData(sheet, range, masterHeaderRowNum);
    Logger.log("objectRowsData[arrayTraderNames[i]][0] = " + objectRowsData[arrayTraderNames[i]][0]);
    
    // filter and keep only rows where all required data is completed
    filteredObjectRowsData[arrayTraderNames[i]] = objectRowsData[arrayTraderNames[i]].filter(function(x) { 
      return x.reportingDataComplete == "Yes"; 
    }); 
    Logger.log("filteredObjectRowsData[arrayTraderNames[i]][0] = " + filteredObjectRowsData[arrayTraderNames[i]][0]);
  }
  return filteredObjectRowsData;
}


function createDetailandSummaryReport(filteredObjectRowsData, settings, d, monthAndYear, isManualReport) {
  Logger.log("**** Running createDetailandSummaryReport() ****");
  /****************************************************************************************************************
  |  First, create data with just columns needed for report
  |  Then, copy data into the detailed report tab of the spreadsheet
  /***************************************************************************************************************/
  var arrayTraderNames = settings.arrayTraderNames;
  var numberOfMasterSheets = settings.arrayTraderNames.length;
  var summaryReportData = [];
  var totalProfitForAllTraders = 0;
  var totalCcFinalInvoiceAmountforAllTraders = 0;
  var totalNumberOfRowsInAllMasterSheets = 0
  
  if (isManualReport) {
    var detailedReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Manual - Detail Report");
    if (detailedReportSheet == null) {
      throw "Error:  The sheet 'Weekly - Detail Report' could not be found";
    }
    var summaryReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Manual - Summary Report");
    if (summaryReportSheet == null) {
      throw "Error:  The sheet 'Weekly - Summary Report' could not be found";
    }
  } else {
    var detailedReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly - Detail Report");
    if (detailedReportSheet == null) {
      throw "Error:  The sheet 'Weekly - Detail Report' could not be found";
    }
    var summaryReportSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Weekly - Summary Report");
    if (summaryReportSheet == null) {
      throw "Error:  The sheet 'Weekly - Summary Report' could not be found";
    }
  }
    
  // Loop through the Master spreadsheets
  for (var i = 0; i < numberOfMasterSheets; i++) {
    var detailedReportData = [];
    var numberOfRowsInMasterSheet = filteredObjectRowsData[arrayTraderNames[i]].length;
    var sumTotalProfit = 0;
    var totalLoads = 0;
    var totalCcFinalInvoiceAmount = 0;
    
    totalNumberOfRowsInAllMasterSheets += numberOfRowsInMasterSheet;
    
    Logger.log("numberOfRowsInMasterSheet = " + numberOfRowsInMasterSheet);
    Logger.log("filteredObjectRowsData[arrayTraderNames[" + i + "]][0] = " + filteredObjectRowsData[arrayTraderNames[i]][0]);
    
    if (numberOfRowsInMasterSheet === 0) {
      continue;
    }
    
    // Check to make sure all needed keys exist
    if (!("traderName" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Trader Name column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("masterRecord" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Master Record # column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("customer" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Customer column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("vendor" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Vendor column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("carrierShipVia" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Carrier / Ship Via column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("ccFinalInvoiceAmount" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find 3CC Final Invoice Amount column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("vendorInvAmountFinal" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Vendor Inv Amount Final column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("brokerFeeInDollars" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Broker Fee in Dollars column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("freightRate" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Freight Rate column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("fsc" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find FSC column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("demurrageDeadFreightWashouts" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Demurrage / Dead Freight / Washouts column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("labFees" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Lab Fees column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("fedexFees" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find FedEx Fees column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("overheadFees" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Overhead Fees column in " + arrayTraderNames[i] + "\'s Master";
    } else if (!("miscFees" in filteredObjectRowsData[arrayTraderNames[i]][0])) {
      throw "Error:  Could not find Misc Fees column in " + arrayTraderNames[i] + "\'s Master";
    }
             
    // loop through the rows in sheet
    for (var j = 0; j < numberOfRowsInMasterSheet; j++) {
      Logger.log('filteredObjectRowsData[arrayTraderNames[i]][j] = ' + filteredObjectRowsData[arrayTraderNames[i]][j]);
      Logger.log('Object keys = ' + Object.keys(filteredObjectRowsData[arrayTraderNames[i]][j]));
      
      if (isNaN(filteredObjectRowsData[arrayTraderNames[i]][j]["ccFinalInvoiceAmount"]) 
          || isNaN(filteredObjectRowsData[arrayTraderNames[i]][j]["vendorInvAmountFinal"])
          || !(isValidExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["brokerFeeInDollars"], "NA"))
          || !(isValidExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["freightRate"], "FOB"))
          || !(isValidExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["fsc"], "NA"))
          || !(isValidExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["demurrageDeadFreightWashouts"], "NA"))
          || !(isValidExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["labFees"], "NA"))
          || !(isValidExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["fedexFees"], "NA"))
          || !(isValidExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["overheadFees"], "NA"))
          || !(isValidExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["miscFees"], "NA")) ) {
        detailedReportData[j] = [];       
        detailedReportData[j].push(d);
        detailedReportData[j].push(monthAndYear);
        detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["traderName"]);
        detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["masterRecord"]);
        detailedReportData[j].push("Skipped");
        continue;
      }
      
      filteredObjectRowsData[arrayTraderNames[i]][j]["brokerFeeInDollars"] = convertExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["brokerFeeInDollars"], "NA");
      filteredObjectRowsData[arrayTraderNames[i]][j]["freightRate"] = convertExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["freightRate"], "FOB");
      filteredObjectRowsData[arrayTraderNames[i]][j]["fsc"] = convertExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["fsc"], "NA");
      filteredObjectRowsData[arrayTraderNames[i]][j]["demurrageDeadFreightWashouts"] = convertExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["demurrageDeadFreightWashouts"], "NA");
      filteredObjectRowsData[arrayTraderNames[i]][j]["labFees"] = convertExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["labFees"], "NA");
      filteredObjectRowsData[arrayTraderNames[i]][j]["fedexFees"] = convertExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["fedexFees"], "NA");
      filteredObjectRowsData[arrayTraderNames[i]][j]["overheadFees"] = convertExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["overheadFees"], "NA");
      filteredObjectRowsData[arrayTraderNames[i]][j]["miscFees"] = convertExpense(filteredObjectRowsData[arrayTraderNames[i]][j]["miscFees"], "NA");

      var totalProfit = filteredObjectRowsData[arrayTraderNames[i]][j]["ccFinalInvoiceAmount"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["vendorInvAmountFinal"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["brokerFeeInDollars"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["freightRate"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["fsc"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["demurrageDeadFreightWashouts"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["labFees"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["fedexFees"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["overheadFees"] - 
                filteredObjectRowsData[arrayTraderNames[i]][j]["miscFees"];   
      Logger.log("totalProfit = " + totalProfit);
      
      // Create data for detailed report
      detailedReportData[j] = [];       
      detailedReportData[j].push(d);
      detailedReportData[j].push(monthAndYear);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["traderName"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["masterRecord"]);
      detailedReportData[j].push(totalProfit);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["customer"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["vendor"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["carrierShipVia"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["ccFinalInvoiceAmount"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["vendorInvAmountFinal"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["brokerFeeInDollars"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["freightRate"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["fsc"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["demurrageDeadFreightWashouts"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["labFees"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["fedexFees"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["overheadFees"]);
      detailedReportData[j].push(filteredObjectRowsData[arrayTraderNames[i]][j]["miscFees"]);
      detailedReportData[j].push(.5);
      detailedReportData[j].push(.5 * (totalProfit - 250));
      
      
      // Calculate totals for summary report
      sumTotalProfit += totalProfit;
      totalCcFinalInvoiceAmount += filteredObjectRowsData[arrayTraderNames[i]][j]["ccFinalInvoiceAmount"];
    }
    // Create data for summary report
    summaryReportData[i] = [];
    summaryReportData[i].push(d);
    summaryReportData[i].push(monthAndYear);
    summaryReportData[i].push(arrayTraderNames[i]);
    summaryReportData[i].push(sumTotalProfit.toFixed(2));
    summaryReportData[i].push((sumTotalProfit / totalCcFinalInvoiceAmount).toFixed(2));
    summaryReportData[i].push(totalCcFinalInvoiceAmount.toFixed(2));
    summaryReportData[i].push((sumTotalProfit / numberOfRowsInMasterSheet).toFixed(2));
    summaryReportData[i].push((totalCcFinalInvoiceAmount / numberOfRowsInMasterSheet).toFixed(2));
    summaryReportData[i].push(numberOfRowsInMasterSheet);
    
    // Calculate totals for email report
    totalProfitForAllTraders += sumTotalProfit;
    totalCcFinalInvoiceAmountforAllTraders += totalCcFinalInvoiceAmount
    
    // Copy data into detail report spreadsheet
    var detailedReportDestinationRange = detailedReportSheet.getRange(detailedReportSheet.getLastRow() + 1, 1, numberOfRowsInMasterSheet, detailedReportData[0].length);   // 1 because data starts in first column
    detailedReportDestinationRange.setValues(detailedReportData);
  }

  
  if (totalNumberOfRowsInAllMasterSheets === 0) {
    return null;
  } else {
    // Copy data into summary report spreadsheet
    Logger.log("summaryReportSheet.getLastRow() = " + summaryReportSheet.getLastRow());
    Logger.log("numberOfMasterSheets = " + numberOfMasterSheets);
    Logger.log("summaryReportData[0].length = " + summaryReportData[0].length);
    var summaryReportDestinationRange = summaryReportSheet.getRange(summaryReportSheet.getLastRow() + 1, 1, numberOfMasterSheets, summaryReportData[0].length);   // 1 because data starts in first column
    summaryReportDestinationRange.setValues(summaryReportData);  
    SpreadsheetApp.flush();
    return {totalProfitForAllTraders:totalProfitForAllTraders,
            totalCcFinalInvoiceAmountforAllTraders:totalCcFinalInvoiceAmountforAllTraders};
  }
}

function test() {
  Logger.log("*** Running test() ***");
  var mySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Test');
  var myData = mySheet.getDataRange().getValues();
  var myRange = mySheet.getRange(2, 1, mySheet.getLastRow()-1, 4);
  myRange.setBackground('#ff0000');
  var rowsData = getRowsData(mySheet, myRange, 0);
  
  var testArray = [];
  for (var i = 0; i < 5; i++) {
    testArray[i] = [];
    testArray[i].push('Date');
    testArray[i].push('MonthAndYear');
  }
  
  var pasteRange = mySheet.getRange(10,1,5,2);
  pasteRange.setValues(testArray);
  
  
  Logger.log("data = " + myData);
  Logger.log("data[0] = " + myData[0]);
  Logger.log("data[0][1] = " + myData[0][1]);
  Logger.log("range = " + myRange);
  Logger.log("rowsData = " + rowsData);
  Logger.log("rowsData[1] = " + rowsData[1]);
  Logger.log("rowsData[1]['name'] = " + rowsData[1]['name']);
  
}

// ******************************************************************** Helper Functions *******************************************************************

function letterToColumn(letter)
{
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++)
  {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function copyTo(source, dest) {
  var sourceSheet = source.getSheet();
  var destSheet = dest.getSheet();
  var sourceData = source.getValues();
  var destRange = destSheet.getRange(
    dest.getRow(),        // Top row of Dest
    dest.getColumn(),     // left col of Dest
    sourceData.length,           // # rows in source
    sourceData[0].length);       // # cols in source (elements in first row)
  destRange.setValues(sourceData);
  // SpreadsheetApp.flush();
}

function getMonthName(d) {
  var month = new Array();
  month[0] = "January";
  month[1] = "February";
  month[2] = "March";
  month[3] = "April";
  month[4] = "May";
  month[5] = "June";
  month[6] = "July";
  month[7] = "August";
  month[8] = "September";
  month[9] = "October";
  month[10] = "November";
  month[11] = "December";
  var m = month[d.getMonth()];
  Logger.log("d.getMonth() = " + d.getMonth());
  return m;
}

function getMonthAndYear(d) {
  var m = getMonthName(d);  // Jan=0, Feb=1, etc
  var y = d.getFullYear(d);
  Logger.log("m = " + m);
  Logger.log("y = " + y);
  
  return m + " " + y;
}

/*function newArrayWithValue(myValue, myLength) {
  var myArray = [];
  for (var i = 0; i < myLength; i++) {
    myArray.push = myValue;
  }
  return myArray;
}*/
  
// ************************* Sheet processing library functions from https://developers.google.com/apps-script/articles/mail_merge#section4 *************

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(range.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

function isNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function isValidExpense(n, compareString) {
  upperN = n.toUpperCase()
  return isNumeric(n) || (upperN == compareString)
}

function convertExpense(n, compareString) {
  upperN = n.toUpperCase()
  if (upperN == compareString) {
    return 0
  } else {
    return n
  }
}

/*
Change minute trigger to trigger once per week
*/
