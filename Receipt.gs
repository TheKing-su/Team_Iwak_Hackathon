var receiptID = "1KcBdZqM2wiMfyAzEAeJ4DXJvzItdeL-CODVi3LqN8OI";

function receiptCheckStatus(invoiceID)
{
  Logger.log(invoiceID);

  var salesDatabaseID = "1rIoeWrQI6NvJategdgNaPrITlr5FV6PrgVfThsVW13s";
  let database = SpreadsheetApp.openById(salesDatabaseID);
  let invoiceData = database.getSheetByName('InvoiceData');

  var invoiceNum = invoiceData.getDataRange().getValues();

  var paymentStatus = "";
  var found = false;

  for (let i = 0; i < invoiceNum.length; i++)
  {
    if(invoiceNum[i][0] === invoiceID)
    {
      found = true;
      paymentStatus = invoiceNum[i][4];
      if(paymentStatus === 'Paid')
      {
        Logger.log('Status: Paid');
        updateReceipt(invoiceNum[i][0]);
        Logger.log(invoiceNum[i][0]);
        return;
      }
      else
      {
        Logger.log('Status: Pending. ');
        return;
      }
    }
  }

  if(!found)
  {
    Logger.log("Invalid Data");
  }
}

function receiptFolder()
{
  const date = new Date();
  const month = date.getMonth();
  const year = date.getFullYear();

  var parentFolderID = "1HtdXCCl-_s07CbxCifMaDQUK5F9tGfOA";
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

function receiptNumber()
{
  const date = new Date();
  const month = date.getMonth()+1;
  const year = date.getFullYear().toString().slice(-2);
  
  var folderID = receiptFolder();
  var folder = DriveApp.getFolderById(folderID);
  var files = folder.getFiles();

  var count = 0;
  while (files.hasNext()) {
    var file = files.next();
    count++;
  }

  Logger.log('Total Receipt: ' + count);

  count++;

  var paddedMonth = ('0' + month).slice(-2);
  var receiptNumber = year + paddedMonth + '_RECEIPT_' + ('000' + count).slice(-3);

  Logger.log("New receipt Number: " + receiptNumber);

  return receiptNumber;
}

function clearReceipt()
{
  let receipt = SpreadsheetApp.openById(receiptID).getSheetByName('Receipt');

  receipt.getRange('B16:B20').clearContent();
  receipt.getRange('C16:F16').clearContent();
  receipt.getRange('C17:F17').clearContent();
  receipt.getRange('C18:F18').clearContent();
  receipt.getRange('C19:F19').clearContent();
  receipt.getRange('C20:F20').clearContent();
}

function updateReceipt(invoiceNum)
{
  clearReceipt();

  const date = new Date();
  Logger.log(date);

  const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  let database = SpreadsheetApp.openById(salesDatabaseID);
  let receiptSheet = SpreadsheetApp.openById(receiptID);
  let invoiceData = database.getSheetByName('InvoiceData');
  let orderDetails = database.getSheetByName('OrderDetails');
  let customerDetails = database.getSheetByName('CustomerData')
  let receipt = receiptSheet.getSheetByName('Receipt');
  let receiptNo = receiptNumber();

  var invoiceCheck = invoiceData.getDataRange().getValues();
  var orderID = "";

  for (let i = 0; i < invoiceCheck.length; i++)
  {
    if(invoiceCheck[i][0] === invoiceNum)
    {
      orderID = invoiceCheck[i][8];
      var dueDate = invoiceCheck[i][3];
      var beforeTax = invoiceCheck[i][5];
      var afterTax = invoiceCheck[i][6];
      break;
    }
  }

  if (!orderID) { 
  Logger.log("Invalid Invoice ID");
  return;
  }

  var orderCheck = orderDetails.getDataRange().getValues();
  var customerID = '';

  var startRow = 16;
  var num = 0;

  for (let j = 0; j < orderCheck.length; j++)
  {
    if(orderCheck[j][0] === orderID)
    {
      customerID = orderCheck[j][7];
      var itemName = orderCheck[j][2];
      var unitPrice = orderCheck[j][3];
      var amount = orderCheck[j][4];
      var total = unitPrice * amount;

      receipt.getRange('B' + startRow).setValue(num+1); //No. 
      receipt.getRange('C' + startRow).setValue(itemName); // Item name
      receipt.getRange('E' + startRow).setValue(amount); // Quantity
      receipt.getRange('F' + startRow).setValue(unitPrice); // Unit Price

      startRow++;
      num++;
      Logger.log(itemName, amount, unitPrice);
    }
  }

  if(!orderID)
  {
    Logger.log("Invalid Order ID");
    return;
  }

  var customerCheck = customerDetails.getDataRange().getValues();
  
  for(let k = 0; k < customerCheck.length; k++)
  {
    if(customerCheck[k][0] === customerID)
    {
      var customerName = customerCheck[k][1];
      var customerAddress = customerCheck[k][2];
      var customerContact = customerCheck[k][3];
      break;
    }
  }

  if (!customerID) { 
  Logger.log("Invalid Customer ID");
  return;
  }

  var sheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName("Form Responses"); 
  var lastRow = sheet.getLastRow();
  var row = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  var customerEmail = row[3]; 

  receipt.getRange('B9:C9').setValue(customerName);
  receipt.getRange('B10:C10').setValue(customerContact);
  receipt.getRange('B11:C11').setValue(customerAddress);
  receipt.getRange('G8').setValue(receiptNo);
  receipt.getRange('G9').setValue(formattedDate);
  receipt.getRange('G10').setValue(orderID);
  receipt.getRange('G11').setValue(dueDate);

  SpreadsheetApp.flush();
  exportReceiptPDF(customerEmail, receiptID, receiptNo);
}

function exportReceiptPDF(customerEmail, receiptID, receiptNo)
{
  let invoicesheet = DriveApp.getFileById(receiptID);

  let blob = invoicesheet.getAs('application/pdf');
  let pdf = DriveApp.getFolderById(receiptFolder())
  .createFile(blob)
  .setName(receiptNo);

  sendEmail(customerEmail, receiptNo, pdf)
}

function sendEmail(customerEmail, receiptName, pdf)
{
  GmailApp.sendEmail(customerEmail, 'Receipt ' + receiptName, 
  'Dear Valued Customer,\n\n' +
  'Thank you for your recent payment. Attached you will find your receipt along with a feedback form. <https://forms.gle/9JeXqFUxjtVP5MmD9>' +
  'We kindly request you to complete the feedback form to help us improve our services.\n\n' +
  'Your satisfaction is our priority, and we appreciate your cooperation.\n\n' +
  'Best regards,\n' +
  '[Your Company Name]', 
  {attachments: pdf});
}