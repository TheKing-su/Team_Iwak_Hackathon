let salesDatabaseID = '1rIoeWrQI6NvJategdgNaPrITlr5FV6PrgVfThsVW13s';
let invoiceID = '1l5CxD3bacQ1K3wDLlCiWDCDKdi6PVIHTA7G-Wozb_uc';
let invoiceSheet = SpreadsheetApp.openById(invoiceID);

function clearInvoice()
{
  let invoice = invoiceSheet.getSheetByName('Invoice');
  invoice.getRange('B9:C11').clearContent();
  invoice.getRange('G9:H9').clearContent();
  invoice.getRange('C21').clearContent();
  invoice.getRange('B16:F20').clearContent();
  invoice.getRange('D9').clearContent();
}

function addItemToDatabase(orderID, customerID, itemID, price, amount)
{
  let databaseid = '1rIoeWrQI6NvJategdgNaPrITlr5FV6PrgVfThsVW13s';
  let databasesheet = SpreadsheetApp.openById(databaseid);
  let orderUpdate = databasesheet.getSheetByName('OrderDetails');
  let items = databasesheet.getSheetByName('ItemDetails');

  var itemdata = items.getDataRange().getValues();
  for(let i = 0; i < itemdata.length; i++)
  {
    if(itemdata[i][0] === itemID)
    {
      var itemname = itemdata[i][1];
      break;
    }
  }

  orderUpdate.appendRow([
      orderID,
      itemID,
      itemname,
      price,
      amount,
      price * amount,
      new Date(),
      customerID
  ])

  Logger.log('Added item to database: OrderID=' + orderID + ', ItemID=' + itemID);
}

function convertData()
{
  let databaseOrderDetails = SpreadsheetApp.openById(salesDatabaseID);
  let orderdetails = databaseOrderDetails.getSheetByName('OrderOverall');

  var lastRow = orderdetails.getLastRow();
  var orderData = orderdetails.getRange(lastRow, 1, 1, orderdetails.getLastColumn()).getValues()[0];

  var orderID = orderData[0];
  var customerID = orderData[2];

  var itemColumns = 
  [
    {addOrder: 6, itemID: 7, price: 8, amount: 9},    // Add Order 2
    {addOrder: 10, itemID: 11, price: 12, amount: 13}, // Add Order 3
    {addOrder: 14, itemID: 15, price: 16, amount: 17}, // Add Order 4
    {addOrder: 18, itemID: 19, price: 20, amount: 21}  // Add Order 5
  ];

  addItemToDatabase(orderID, customerID, orderData[3], orderData[4], orderData[5]);
  updateQuantity(orderData[3], orderData[5]);

  for (var i = 0; i < itemColumns.length; i++) 
  {
    if (orderData[itemColumns[i].addOrder] == 'Add' &&
        orderData[itemColumns[i].itemID] &&
        orderData[itemColumns[i].price] &&
        orderData[itemColumns[i].amount]) {
      
      addItemToDatabase(
        orderID,
        customerID,
        orderData[itemColumns[i].itemID],
        orderData[itemColumns[i].price],
        orderData[itemColumns[i].amount]
      );

      updateQuantity(       
        orderData[itemColumns[i].itemID],
        orderData[itemColumns[i].amount]);
    }
  }

  Logger.log('Order ID: ' + orderID);
  Logger.log('Customer ID: ' + customerID);
  Logger.log('Processing item columns...');

  
  invoice(orderID);
}

function GenerateNewID(invoiceData){
  var data = invoiceData.getRange('A:A').getValues();

  var lastrow = null;

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== null && data[i][0] !== undefined && data[i][0] !== '') {
      lastrow = data[i][0];
      break;
    }
  }

  var numpart = parseInt(lastrow.slice(1));

  if (isNaN(numpart)) {
    numpart = 0;
  }

  var nextnum = numpart + 1;
  var nextid = 'I' + ('00' + nextnum);

  return nextid;

}

function updateInvoiceDatabase(invoiceNumber, dueDate, beforeTax, afterTax, testData)
{
  var salesDatabase = SpreadsheetApp.openById(salesDatabaseID);
  var invoiceData = salesDatabase.getSheetByName('InvoiceData');
  var lastrow = GenerateNewID(invoiceData);

  Logger.log(lastrow);

  var newData = [
    lastrow,
    invoiceNumber,
    new Date(), 
    dueDate,
    'Pending',
    beforeTax,
    afterTax,
    null,
    testData
  ];

  invoiceData.appendRow(newData);
  
}

function invoiceFolder()
{
  const date = new Date();
  const month = date.getMonth();
  const year = date.getFullYear();

  var parentFolderID = '1sRumCNxTYEVKPzBexdrsSIUPoX73UaG6';
  const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const folderName = monthNames[month] + " " + year;

  const parentFolder = DriveApp.getFolderById(parentFolderID);
  const folders = parentFolder.getFoldersByName(folderName);

  let invoiceFolder;

  if(folders.hasNext())
  {
    invoiceFolder = folders.next();
    Logger.log("Opened existing folder: " + invoiceFolder.getName());
  } 
  else 
  {
    invoiceFolder = parentFolder.createFolder(folderName);
    Logger.log("Created new folder: " + invoiceFolder.getName());
  }
  
  return invoiceFolder.getId();
}

function invoiceNum()
{
  const date = new Date();
  const year = date.getFullYear();
  
  var folderID = invoiceFolder();
  var folder = DriveApp.getFolderById(folderID);
  var files = folder.getFiles();

  var count = 0;
  while (files.hasNext()) {
    var file = files.next();
    count++;
  }

  Logger.log('Total invoice: ' + count);

  count++;

  var numcount = ('000000' + count).slice(-6);
  var invoiceNumber = year + 'INV' + numcount;

  Logger.log("New Invoice Number: " + invoiceNumber);

  return invoiceNumber;
}

function exportInvoicePDF(customerEmail, invoiceNumber)
{
    let invoicesheet = DriveApp.getFileById(invoiceID);
    let blob = invoicesheet.getAs('application/pdf');

    let folderID = invoiceFolder();
    let folder = DriveApp.getFolderById(folderID);

    let pdf = folder.createFile(blob);
    pdf.setName(invoiceNumber);

    sendinvoiceEmail(customerEmail, invoiceNumber, pdf);
}

function invoice(orderID)
{
  let database = SpreadsheetApp.openById(salesDatabaseID);
  let invoiceSheet = SpreadsheetApp.openById(invoiceID);
  let orderDetails = database.getSheetByName('OrderDetails');
  let customerDetails = database.getSheetByName('CustomerData')
  let invoice = invoiceSheet.getSheetByName('Invoice');
  let invoiceNumber = invoiceNum();

  var orderCheck = orderDetails.getDataRange().getValues();
  var customerID = '';

  var startRow = 16;
  var beforeTax = 0;

  for (let j = 0; j < orderCheck.length; j++)
  {
    if(orderCheck[j][0] === orderID)
    {
      customerID = orderCheck[j][7];
      var itemName = orderCheck[j][2];
      var unitPrice = orderCheck[j][3];
      var amount = orderCheck[j][4];
      var total = unitPrice * amount;

      beforeTax += total;

      invoice.getRange('B' + startRow).setValue(itemName); // Item name
      invoice.getRange('E' + startRow).setValue(amount); // Quantity
      invoice.getRange('F' + startRow).setValue(unitPrice); // Unit Price

      startRow++;
      Logger.log(itemName + amount + unitPrice);
    }
  }

  var afterTax = beforeTax * 1.16;

  if(!customerID)
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
      var customerEmail = customerCheck[k][4];
      var reliabilityStatus = customerCheck[k][5];
      break;
    }
  }

  if (!customerID) { 
  Logger.log("Invalid Customer ID");
  return;
  }

  const todayDate = new Date();
  const formattedDate = Utilities.formatDate(todayDate, Session.getScriptTimeZone(), 'd/M/yyyy');
  var dueDate = new Date(todayDate);

  switch (reliabilityStatus){
    case "Low":
      dueDate.setDate(todayDate.getDate() + 30);
      break;
    case "Medium":
      dueDate.setDate(todayDate.getDate() + 60);
      break;
    case "High":
      dueDate.setDate(todayDate.getDate() + 90);
      break;
  }

  invoice.getRange('G9').setValue(invoiceNumber);
  invoice.getRange('B13').setValue(formattedDate);
  invoice.getRange('C21').setValue(dueDate);
  invoice.getRange('B9:C9').setValue(customerName +' ('+  customerID + ')');
  invoice.getRange('B10:C10').setValue(customerAddress);
  invoice.getRange('B11:C11').setValue(customerContact);

  SpreadsheetApp.flush();
  updateInvoiceDatabase(invoiceNumber, dueDate, beforeTax, afterTax, orderID); 
  exportInvoicePDF(customerEmail, invoiceNumber)
}

function sendinvoiceEmail(customerEmail, invoiceNumber, pdf)
{
  GmailApp.sendEmail(customerEmail, 'Invoice ' + invoiceNumber, 
  'Dear Valued Customer,\n\n' +
  'This is Please find attached your invoice for your recent purchase and the payment form here <https://forms.gle/1k2daSL26VPW1QRH9> \n\n' +
  '[Your Company Name]', 
  {attachments: [pdf]});
}

function extractQuantity() {
  let databaseOrderDetails = SpreadsheetApp.openById(salesDatabaseID);
  let orderdetails = databaseOrderDetails.getSheetByName('OrderOverall');

  var lastRow = orderdetails.getLastRow();
  var orderData = orderdetails.getRange(lastRow, 1, 1, orderdetails.getLastColumn()).getValues()[0];

  var orderID = orderData[0];
  var customerID = orderData[2];

  var itemColumns = [
    {addOrder: 6, itemID: 7, price: 8, amount: 9},    // Add Order 2
    {addOrder: 10, itemID: 11, price: 12, amount: 13}, // Add Order 3
    {addOrder: 14, itemID: 15, price: 16, amount: 17}, // Add Order 4
    {addOrder: 18, itemID: 19, price: 20, amount: 21}  // Add Order 5
  ];

  updateQuantity(orderData[3], orderData[5]);
  
  for (var i = 0; i < itemColumns.length; i++) {
    if (orderData[itemColumns[i].addOrder] == 'Add' &&
        orderData[itemColumns[i].itemID] &&
        orderData[itemColumns[i].amount]) {
      
      updateQuantity(
        orderData[itemColumns[i].itemID],
        orderData[itemColumns[i].amount]
      );
    }
  }
}

function updateQuantity(itemID, amount) {
  let databasesheet = SpreadsheetApp.openById(salesDatabaseID);
  let items = databasesheet.getSheetByName('ItemDetails');

  var itemdata = items.getDataRange().getValues();
  for (let i = 0; i < itemdata.length; i++) {
    if (itemdata[i][0] === itemID) {
      var itemQuantity = itemdata[i][5] - amount;
      items.getRange(i + 1, 6).setValue(itemQuantity);
      Logger.log('Item Quantity Updated. ItemID: ' + itemID + ', NewQuantity=' + itemQuantity);
      break;
    }
  }
}