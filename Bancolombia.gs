// REF: http://www.grupobancolombia.com/contenidoCentralizado/corporativo/formatospdf/SVE/FormatoPagosPAB.pdf

function CreatePaymentFile() {  
  var activeSheet =  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  // This represents ALL the data
  var range = activeSheet.getRange("A10:D100");
  var values = range.getValues();
  
  var fileBancolombia = "";
  
  for (var i = 0; i < values.length; i++) {
    //Valor fijo 6 Indica el tipo de registro de detalle
    var bancolombiaLine = "6";
    for (var j = 0; j < values[i].length; j++) {
      // Get Bank Account Number, Bank, Account Type & Account Name
      if (j==0)
      {
        
      }
      // Get Reference
      if (j==1)
      {
        
      }      
      // Get Amount
      if (j==2)
      {
        
      }       
      // Get Date
      if (j==3)
      {
        
      }         
    }
    fileBancolombia = fileBancolombia + "\r\n";
  }  
  
  // Control Line    

  var paymentFolder = DriveApp.createFolder("PagosBancolombia");
  var tempFile = paymentFolder.createFile("Bancolombia_" + activeSheet.getName() + ".txt", fileBancolombia, "text/plain");
}
