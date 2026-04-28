

function summarizeCounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var summarySheetName = "Summary";

  // Mapping of sheet names to display names
  var nameMapping = {
    "ACArcoviaCity": "AC - Arcovia City",
    "AWAlabangWest": "AW - Alabang West",
    "Bacolod": "Bacolod",
    "BCBroadwayCentrum": "BC - Broadway Centrum",
    "BNBoracayNewcoast": "BN - Boracay Newcoast",
    "Davao": "Davao",
    "CCFClarkCityfront": "CCF - Clark Cityfront",
    "CGSCaliforniaGardenSquare": "CGS - California Garden Square",
    "EWCEastwoodCity": "EWC - Eastwood City",
    "FTForbesTown": "FT - Forbes Town",
    "IBPIloiloBusinessPark": "IBP - Iloilo Business Park",
    "LCTLuckyChinatown": "LCT - Lucky Chinatown",
    "MGMapleGrove": "MG - Maple Grove",
    "MKHMcKinleyHill": "MKH - McKinley Hill",
    "NPCNewportCity": "NPC - Newport City",
    "MAKATIPROPERTIES": "MAKATI PROPERTIES",
    "SLPSanLorenzo": "SLP - San Lorenzo",
    "SWSouthwoods": "SW - Southwoods",
    "CTDTheClubhouseTempleDrive": "CTD - The Clubhouse Temple Drive",
    "TLTwinLakes": "TL - Twin Lakes",
    "TMNTheMactanNewtown": "TMN - The Mactan Newtown",
    "VSVillageSquare": "VS - Village Square",
    "UBUptownBonifacio": "UB - Uptown Bonifacio"
  };

  // Check if summary sheet exists, if not, create it
  var summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(summarySheetName);
  }

  // Define months for header
  var months = ["January", "February", "March", "April", "May", "June", 
                "July", "August", "September", "October", "November", "December"];
  
  // Construct header row with alternating B-K, T-Q, and Validation count beside T-Q
  var headerRow = ["Property Tab"];
  months.forEach(month => {
    headerRow.push(month + " Total Deficient", month + " Total Rectification", month + " Total Validation");
  });

  var startRow = 31; // Start summary at row 31
  var startCol = 1;

  // Write header only if it's not already there
  var existingHeaders = summarySheet.getRange(startRow, startCol, 1, headerRow.length).getValues()[0];
  if (existingHeaders.join("") !== headerRow.join("")) {
    summarySheet.getRange(startRow, startCol, 1, headerRow.length).setValues([headerRow]);
  }

  var summaryData = [];
  
  // Function to convert various date formats to a Date object
  function parseDate(dateValue) {
    if (dateValue instanceof Date) {
      return dateValue;
    } else if (typeof dateValue === "string") {
      var parsedDate = new Date(dateValue);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate;
      }

      // Handle formats like "February 09, 2025 at 12:59:09"
      var regex = /([A-Za-z]+) (\d{1,2}), (\d{4})/;
      var match = dateValue.match(regex);
      if (match) {
        var monthNames = {
          "January": 0, "February": 1, "March": 2, "April": 3, "May": 4, "June": 5,
          "July": 6, "August": 7, "September": 8, "October": 9, "November": 10, "December": 11
        };
        var month = monthNames[match[1]];
        var day = parseInt(match[2], 10);
        var year = parseInt(match[3], 10);
        return new Date(year, month, day);
      }
    }
    return null;
  }

  // Loop through each sheet and collect summary data
  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    if (sheetName === summarySheetName) return; // Skip summary sheet

    var displayName = nameMapping[sheetName] || sheetName; // Apply name change if in mapping

    var data = sheet.getDataRange().getValues();
    var monthlyCounts = Array(36).fill(0); // 12 months * 3 (Deficiency, Rectification, Validation)

    for (var i = 1; i < data.length; i++) { // Skip header row
      var bDate = parseDate(data[i][1]); // Column B (index 1)
      var kValue = data[i][10]; // Column K (index 10)

      if (bDate) {
        var monthIndex = bDate.getMonth() * 3; // Adjusted for 3 columns per month
        if (kValue !== "" && kValue !== null) {
          monthlyCounts[monthIndex]++;
        }
      }

      var tDate = parseDate(data[i][19]); // Column T (index 19)
      var qValue = data[i][16]; // Column Q (index 16)
      var vValue = data[i][24]; // Column V (index 21)

      if (tDate) {
        var monthIndex = (tDate.getMonth() * 3) + 1; // Adjusted for 3 columns per month
        if (qValue !== "" && qValue !== null) {
          monthlyCounts[monthIndex]++;
        }

        if (vValue === true || vValue === "TRUE") { // Count "true" values in V
          monthlyCounts[monthIndex + 1]++;
        }
      }
    }

    summaryData.push([displayName, ...monthlyCounts]); // Use mapped name
  });

  var numRows = summaryData.length;
  var numCols = summaryData[0].length;

  // Add column totals
  if (numRows > 0) {
    var totalRow = Array(numCols).fill(0);
    totalRow[0] = "TOTAL"; // Label the row

    for (var col = 1; col < numCols; col++) {
      totalRow[col] = summaryData.reduce((sum, row) => sum + (row[col] || 0), 0);
    }

    // Append total row at the bottom of the summary data
    summaryData.push(totalRow);

    // Overwrite only the necessary cells, keeping other data intact
    summarySheet.getRange(startRow + 1, startCol, summaryData.length, numCols).setValues(summaryData);
  }

  SpreadsheetApp.getUi().alert("Monthly Summary updated in 'Summary' tab!");
}

function onOpen() {
  // Run the functions automatically when the spreadsheet is opened
  summarizeColBTrimmedCounts();
  summarizeCounts();
}


function summarizeColBTrimmedCounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var summarySheet = ss.getSheetByName("Summary") || ss.insertSheet("Summary");

  // Set the starting row for appending data (row 200)
  var startRow = 60;

  // Prepare the headers only if they don't already exist at row 200
  var headers = ["Property", "Day Inspected", "Deficient Count", "Day Rectified", "Rectify Count"];
  var existingHeaders = summarySheet.getRange(startRow, 1, 1, headers.length).getValues()[0];
  
  // If headers are not already there, insert them
  if (existingHeaders.join("") !== headers.join("")) {
    summarySheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
  }

  var allData = [];  // This will store all the summary data to be written at once

  // Process each sheet
  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    if (sheetName === "Summary") return; // Skip summary sheet

    var data = sheet.getDataRange().getValues();
    var bCountMap = {}; // To store counts for Column B
    var tCountMap = {}; // To store counts for Column T

    // Use a single loop to process both Column B and Column T (index 1 and 19)
    for (var i = 1; i < data.length; i++) {
      // Process Column B (index 1)
      processDateColumn(data[i][1], bCountMap);

      // Process Column T (index 19)
      processDateColumn(data[i][19], tCountMap);
    }

    // Prepare and store the results for this sheet
    for (var bDateStr in bCountMap) {
      var tDateStr = tCountMap[bDateStr] ? bDateStr : ""; // Matching date in Column T if exists
      allData.push([sheetName, bDateStr, bCountMap[bDateStr], tDateStr, tCountMap[bDateStr] || 0]);
    }

    // Also store the remaining dates from Column T if not already included
    for (var tDateStr in tCountMap) {
      if (!bCountMap[tDateStr]) { // If the date from Column T is not already added
        allData.push([sheetName, "", 0, tDateStr, tCountMap[tDateStr]]);
      }
    }
  });

  // Write all results at once to avoid multiple appendRow calls
  if (allData.length > 0) {
    summarySheet.getRange(startRow + 1, 1, allData.length, 5).setValues(allData);
  }

  // Show completion message
  SpreadsheetApp.getUi().alert("Day to Day Updated!");

}

// Helper function to process date columns (B or T)
function processDateColumn(rawDate, countMap) {
  if (!rawDate) return;

  try {
    var dateObj = new Date(rawDate);
    // If the date is invalid, clean it up
    if (isNaN(dateObj)) {
      var cleaned = rawDate.toString().split(" at ")[0]; // Clean "March 27, 2025 at 20:43:59"
      dateObj = new Date(cleaned);
    }

    if (isNaN(dateObj)) return; // Skip invalid dates

    var formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MMM dd yyyy");

    if (!countMap[formattedDate]) {
      countMap[formattedDate] = 0;
    }
    countMap[formattedDate]++;
  } catch (e) {
    // Ignore any error and skip malformed data
    Logger.log('Error processing date: ' + rawDate);
  }
}


function summarizeCounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var summarySheetName = "Summary";

  // Mapping of sheet names to display names
  var nameMapping = {
    "ACArcoviaCity": "AC - Arcovia City",
    "AWAlabangWest": "AW - Alabang West",
    "Bacolod": "Bacolod",
    "BCBroadwayCentrum": "BC - Broadway Centrum",
    "BNBoracayNewcoast": "BN - Boracay Newcoast",
    "Davao": "Davao",
    "CCFClarkCityfront": "CCF - Clark Cityfront",
    "CGSCaliforniaGardenSquare": "CGS - California Garden Square",
    "EWCEastwoodCity": "EWC - Eastwood City",
    "FTForbesTown": "FT - Forbes Town",
    "IBPIloiloBusinessPark": "IBP - Iloilo Business Park",
    "LCTLuckyChinatown": "LCT - Lucky Chinatown",
    "MGMapleGrove": "MG - Maple Grove",
    "MKHMcKinleyHill": "MKH - McKinley Hill",
    "NPCNewportCity": "NPC - Newport City",
    "MAKATIPROPERTIES": "MAKATI PROPERTIES",
    "SLPSanLorenzo": "SLP - San Lorenzo",
    "SWSouthwoods": "SW - Southwoods",
    "CTDTheClubhouseTempleDrive": "CTD - The Clubhouse Temple Drive",
    "TLTwinLakes": "TL - Twin Lakes",
    "TMNTheMactanNewtown": "TMN - The Mactan Newtown",
    "VSVillageSquare": "VS - Village Square",
    "UBUptownBonifacio": "UB - Uptown Bonifacio"
  };

  // Check if summary sheet exists, if not, create it
  var summarySheet = ss.getSheetByName(summarySheetName);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(summarySheetName);
  }

  // Define months for header
  var months = ["January", "February", "March", "April", "May", "June", 
                "July", "August", "September", "October", "November", "December"];
  
  // Construct header row with alternating B-K, T-Q, and Validation count beside T-Q
  var headerRow = ["Property Tab"];
  months.forEach(month => {
    headerRow.push(month + " Total Deficient", month + " Total Rectification", month + " Total Validation");
  });

  var startRow = 30; // Start summary at row 31
  var startCol = 1;

  // Write header only if it's not already there
  var existingHeaders = summarySheet.getRange(startRow, startCol, 1, headerRow.length).getValues()[0];
  if (existingHeaders.join("") !== headerRow.join("")) {
    summarySheet.getRange(startRow, startCol, 1, headerRow.length).setValues([headerRow]);
  }

  var summaryData = [];
  
  // Function to convert various date formats to a Date object
  function parseDate(dateValue) {
    if (dateValue instanceof Date) {
      return dateValue;
    } else if (typeof dateValue === "string") {
      var parsedDate = new Date(dateValue);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate;
      }

      // Handle formats like "February 09, 2025 at 12:59:09"
      var regex = /([A-Za-z]+) (\d{1,2}), (\d{4})/;
      var match = dateValue.match(regex);
      if (match) {
        var monthNames = {
          "January": 0, "February": 1, "March": 2, "April": 3, "May": 4, "June": 5,
          "July": 6, "August": 7, "September": 8, "October": 9, "November": 10, "December": 11
        };
        var month = monthNames[match[1]];
        var day = parseInt(match[2], 10);
        var year = parseInt(match[3], 10);
        return new Date(year, month, day);
      }
    }
    return null;
  }

  // Loop through each sheet and collect summary data
  sheets.forEach(sheet => {
    var sheetName = sheet.getName();
    if (sheetName === summarySheetName) return; // Skip summary sheet

    var displayName = nameMapping[sheetName] || sheetName; // Apply name change if in mapping

    var data = sheet.getDataRange().getValues();
    var monthlyCounts = Array(36).fill(0); // 12 months * 3 (Deficiency, Rectification, Validation)

    for (var i = 1; i < data.length; i++) { // Skip header row
      var bDate = parseDate(data[i][1]); // Column B (index 1)
      var kValue = data[i][10]; // Column K (index 10)

      if (bDate) {
        var monthIndex = bDate.getMonth() * 3; // Adjusted for 3 columns per month
        if (kValue !== "" && kValue !== null) {
          monthlyCounts[monthIndex]++;
        }
      }

      var tDate = parseDate(data[i][19]); // Column T (index 19)
      var qValue = data[i][16]; // Column Q (index 16)
      var vValue = data[i][24]; // Column V (index 21)

      if (tDate) {
        var monthIndex = (tDate.getMonth() * 3) + 1; // Adjusted for 3 columns per month
        if (qValue !== "" && qValue !== null) {
          monthlyCounts[monthIndex]++;
        }

        if (vValue === true || vValue === "TRUE") { // Count "true" values in V
          monthlyCounts[monthIndex + 1]++;
        }
      }
    }

    summaryData.push([displayName, ...monthlyCounts]); // Use mapped name
  });

  var numRows = summaryData.length;
  var numCols = summaryData[0].length;

  // Add column totals
  if (numRows > 0) {
    var totalRow = Array(numCols).fill(0);
    totalRow[0] = "TOTAL"; // Label the row

    for (var col = 1; col < numCols; col++) {
      totalRow[col] = summaryData.reduce((sum, row) => sum + (row[col] || 0), 0);
    }

    // Append total row at the bottom of the summary data
    summaryData.push(totalRow);

    // Overwrite only the necessary cells, keeping other data intact
    summarySheet.getRange(startRow + 1, startCol, summaryData.length, numCols).setValues(summaryData);
  }

  SpreadsheetApp.getUi().alert("Monthly Summary updated in 'Summary' tab!");
}

