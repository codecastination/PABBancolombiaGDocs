function CreatePaymentFile() {  
  var activeSheet =  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  // This represents ALL the data
  var range = activeSheet.getRange("A10:D100");
  var values = range.getValues();
  
  var fileBancolombia = "";
  
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j]) {
        fileBancolombia = fileBancolombia + "," + values[i][j];
      }
      fileBancolombia = fileBancolombia + "\r\n";
    }
    
  }  
  
  // Control Line    

    
  var tempFile = DriveApp.createFile("Bancolombia_" + activeSheet.getName() + ".txt", fileBancolombia, "text/plain");
}
