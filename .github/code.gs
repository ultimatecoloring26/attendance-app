function doGet(e) {
  // 1. If someone visits the URL directly, show the HTML App
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile('Index')
        .setTitle("Nik's Attendance Tracker")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // 2. If the App is sending data, process it here
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Employee_Attendance_App");
  if (!sheet) {
     sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }
  
  var action = e.parameter.action; 
  var name = e.parameter.name;
  var email = e.parameter.email;
  var location = e.parameter.location;
  
  var timestamp = new Date();
  var timeString = timestamp.toLocaleTimeString(); // e.g. "8:30:00 AM"
  var dateString = timestamp.toLocaleDateString(); // e.g. "10/4/2026"

 if (action === 'Clock In') {
    // 1. Find the next empty row by looking only at Column A
    var colA = sheet.getRange("A1:A").getValues();
    var newRow = colA.filter(String).length + 1;
    
    // 2. Drop the data into the exact columns (Leaving H alone!)
    sheet.getRange(newRow, 1).setValue(name);       // Column A: Name
    sheet.getRange(newRow, 2).setValue(email);      // Column B: Email
    sheet.getRange(newRow, 3).setValue(dateString); // Column C: Date
    sheet.getRange(newRow, 4).setValue(timeString); // Column D: Time In
    sheet.getRange(newRow, 7).setValue("In: " + location); // Column G: GPS
    
    return ContentService.createTextOutput("✅ Success: Clock In recorded for " + name);
    
  } else {
    // For Clock Out, Start Rest, End Rest -> We must find their open shift for the day
    var dataRange = sheet.getDataRange().getValues();
    
    // Search backwards to find their latest Clock In that doesn't have a Clock Out yet
    for (var i = dataRange.length - 1; i >= 1; i--) {
      
      // dataRange[i][0] is Name (Col A), dataRange[i][4] is Time Out (Col E)
      if (dataRange[i][0] == name && dataRange[i][4] == "") {
        
        var rowNum = i + 1; // Translate array index to Google Sheet row number
        var existingLoc = dataRange[i][6] || ""; // Get existing GPS from Column G
        
        if (action === 'Start Rest') {
          sheet.getRange(rowNum, 9).setValue(timeString); // Column I
          sheet.getRange(rowNum, 7).setValue(existingLoc + " | Rest: " + location);
          return ContentService.createTextOutput("☕ Success: Start Rest recorded!");
        } 
        else if (action === 'End Rest') {
          sheet.getRange(rowNum, 10).setValue(timeString); // Column J
          sheet.getRange(rowNum, 7).setValue(existingLoc + " | Back: " + location);
          return ContentService.createTextOutput("💪 Success: End Rest recorded!");
        } 
        else if (action === 'Clock Out') {
          sheet.getRange(rowNum, 5).setValue(timeString); // Column E
          sheet.getRange(rowNum, 7).setValue(existingLoc + " | Out: " + location);
          return ContentService.createTextOutput("👋 Success: Clock Out recorded!");
        }
      }
    }
    return ContentService.createTextOutput("❌ Error: We couldn't find an open 'Clock In' for you today.");
  }
}
