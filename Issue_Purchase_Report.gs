function purchaseOrderFolder()
{
  const date = new Date();
  const month = date.getMonth();
  const year = date.getFullYear();

  var parentFolderID = "1LMfCo1YqvMd9IWvqgKLu8TrmQrEv7DvV";
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const folderName = monthNames[month] + " " + year;

  const parentFolder = DriveApp.getFolderById(parentFolderID);
  const folders = parentFolder.getFoldersByName(folderName);

  let receiptFolder;

  if(folders.hasNext())
  {
    receiptFolder = folders.next();
    Logger.log("Opened existing folder: " + receiptFolder.getName());
  } 
  else 
  {
    receiptFolder = parentFolder.createFolder(folderName);
    Logger.log("Created new folder: " + receiptFolder.getName());
  }
  
  return receiptFolder.getId();
}

function issuePurchaseOrder()
{
  var sheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName("PurchaseRecord"); 
  var lastRow = sheet.getLastRow();
  var row = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const todayDate = new Date();
  const formattedDate = Utilities.formatDate(todayDate, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  var purchaseID = row[0];
  var itemID = row[1];
  var unitprice = row[2];
  var quantity = row[3];
  var supplierID = row[5];

  Logger.log(supplierID);

  var suppliersheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName('SupplierData');
  var supplierData = suppliersheet.getDataRange().getValues();

  for(let i = 0; i < supplierData.length; i++)
  {
    if(supplierData[i][0] === supplierID)
    {
      var supplierName = supplierData[i][1];
      var supplierAddress = supplierData[i][2];
      var supplierContact = supplierData[i][3];
      var supplierEmail = supplierData[i][4];
      var supplierShipping = supplierData[i][5]
      var supplierHandling = supplierData[i][6];
      var supplierStorage = supplierData[i][7];
      continue;
    }
  }

  Logger.log(supplierEmail);

  var othercost = supplierShipping + supplierHandling + supplierStorage;

  var itemsheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName('ItemDetails');
  var itemData = itemsheet.getDataRange().getValues();

  for(let j = 0; j < itemData.length; j++)
  {
    if(itemData[j][0] === itemID)
    {
      var itemName = itemData[j][1];
      var itemQuantity = itemData[j][5];
      var index = j;
      break;
    }
  }
  
  let purchaseorderID = '1vwVrxpzglR4WGjLy-i5EQwSz1ZRdZ64M0SkeKw---XQ';
  let purchasesheet = SpreadsheetApp.openById(purchaseorderID).getSheetByName('Purchase order');

  purchasesheet.getRange('B11').setValue(formattedDate);
  purchasesheet.getRange('D11:E11').setValue(supplierName);
  purchasesheet.getRange('F11:G11').setValue(purchaseID);
  purchasesheet.getRange('B17:C17').setValue(supplierName);
  purchasesheet.getRange('B18:C19').setValue(supplierAddress);
  purchasesheet.getRange('B20').setValue(supplierContact);
  purchasesheet.getRange('B24').setValue(itemID);
  purchasesheet.getRange('C24').setValue(itemName);
  purchasesheet.getRange('F24').setValue(quantity);
  purchasesheet.getRange('G24').setValue(unitprice);
  purchasesheet.getRange('H34').setValue(othercost);


  let pdfFolder = DriveApp.getFolderById(purchaseOrderFolder());
  let fileName = 'PurchaseOrder_' + purchaseID + '.pdf';
  let blob = purchasesheet.getParent().getAs('application/pdf').setName(fileName);
  let pdfFile = pdfFolder.createFile(blob);

  var subject = 'Purchase Order: ' + purchaseID;
  var body = 'Dear ' + supplierName + ',\n\n' +
             'Please find attached the purchase order for the following items:\n\n' +
             '- Item ID: ' + itemID + '\n' +
             '- Item Name: ' + itemName + '\n' +
             '- Quantity: ' + quantity + '\n' +
             '- Unit Price: RM' + unitprice + '\n' +
             '\nTotal Additional Costs: RM' + othercost + '\n\n' +
             'Best regards,\nYour Company Name';

  MailApp.sendEmail({
    to: supplierEmail,
    subject: subject,
    body: body,
    attachments: [pdfFile.getAs(MimeType.PDF)]
  })

  Logger.log('Purchase order sent to ' + supplierEmail);
  costCalculation(supplierID);
}