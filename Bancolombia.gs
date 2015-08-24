// REF: http://www.grupobancolombia.com/contenidoCentralizado/corporativo/formatospdf/SVE/FormatoPagosPAB.pdf

function CreatePaymentFile() {  
  var activeSheet =  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
 
  // Payroll Data
  var payrollData = activeSheet.getRange("A10:D100").getValues();
  
  // 3rd Party Data
  var DB3rd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TERCEROS").getRange("A2:E100").getValues();
  // 3rd Party Data
  var DBBanks = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BANCOS").getRange("A2:B100").getValues();
  // 3rd Party Data
  var DBAccType = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TIPO CUENTA").getRange("A2:E100").getValues();  
  
  var fileBancolombia = "";
  
  for (var i = 0; i < payrollData.length; i++) {
    //Valor fijo 6 Indica el tipo de registro de detalle
    var bancolombiaLine = "6";
    var accountName = payrollData[i][0];
    var amount = payrollData[i][1];
    var txnDate = payrollData[i][2];
    var reference = payrollData[i][3];
    if (accountName.length > 0){
      // Get Bank Account Number, Bank, Account Type & Account Name
      bancolombiaLine += GetDocumentNumber(DB3rd, accountName);
      bancolombiaLine += GetBankAccountName(accountName);
      bancolombiaLine += GetBankAccount(DB3rd, DBBanks, DBAccType, accountName); 
      // Get Amount
      bancolombiaLine += GetAmount(amount);     
      // Get Date
      bancolombiaLine += GetTxnDate(txnDate);   
      // Get Reference
      bancolombiaLine += GetReference(reference);  
      // Get Final Part
      bancolombiaLine += GetFinalPart();
      
      fileBancolombia += bancolombiaLine + "\r\n";
    }
  }  
  
  // Folder 
  var folders = DriveApp.getFoldersByName("PagosBancolombia");
  var bancolombiaFolder = null;
  while(folders.hasNext()){
    var folderFound = folders.next();
    if(folderFound.getName() == "PagosBancolombia"){
      bancolombiaFolder = folderFound;
    }
  }
  if (bancolombiaFolder == null){
    bancolombiaFolder = DriveApp.createFolder("PagosBancolombia");
  }
  var tempFile = bancolombiaFolder.createFile("Bancolombia_" + activeSheet.getName() + ".txt", fileBancolombia, "text/plain");
}

function GetFinalPart(){
  return "000000                                                                                                                                         ";
}

function GetDocumentNumber(DB, accountName)
{
  for (var i = 0; i < DB.length; i++) {
    var entryValue = DB[i][1]; 
    if (entryValue == accountName){
      return padRight(DB[i][0].toString(), " ", 15);
    }
  }
}

function GetBankAccount(DB3rd, DBBanks, DBAccType, accountName)
{
  var accountRow;
  var accountBankRow;
  var accountTypeRow;
  // find account row
  for (var i = 0; i < DB3rd.length; i++) {
    var entryValue = DB3rd[i][1]; 
    if (entryValue == accountName){
      accountRow = DB3rd[i];
    }
  }
  // find account bank
  for (var i = 0; i < DBBanks.length; i++) {
    var entryValue = DBBanks[i][0]; 
    if (entryValue == accountRow[2]){
      accountBankRow = DBBanks[i];
    }
  }  
  // find account type
  for (var i = 0; i < DBAccType.length; i++) {
    var entryValue = DBAccType[i][0]; 
    if (entryValue == accountRow[3]){
      accountTypeRow = DBAccType[i];
    }
  }   
  //     Account Id                   + Account Number                              + Account Type
  return accountBankRow[1].toString() + padRight(accountRow[4].toString(), " ", 17) + accountTypeRow[1].toString();
}

function GetBankAccountName(accountName){
  return padRight(accountName, " ", 30);
}
  
function GetAmount(amount)
{
  var amt = Number(amount).toFixed(2).toString().replace(".","").replace(",","");
  return ("00000000000000000" + amt).slice(-17);
}  

function GetTxnDate(txnDate)
{
  txnDate = new Date(txnDate);
  var pad = "00";
  // year
  var retDate = txnDate.getFullYear();
  // month
  retDate += ("00" + (txnDate.getMonth()+1)).slice(-2);
  // day
  retDate += ("00" + txnDate.getDate()).slice(-2);
  
  return retDate;
}

function GetReference(reference)
{
  return padRight(reference, " ", 21);
}

function padRight(s, c, n) {
  if (! s || ! c || s.length >= n) {
    return s;
  }
  var max = (n - s.length)/c.length;
  for (var i = 0; i < max; i++) {
    s += c;
  }
  return s;
}
