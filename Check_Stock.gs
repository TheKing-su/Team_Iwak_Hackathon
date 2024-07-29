function checkStockLevels() {
  const sheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName('ItemDetails');
  const data = sheet.getDataRange().getValues(); 
  var status = ''

  for (let i = 1; i < data.length; i++) {
    const itemID = data[i][0];
    const itemName = data[i][1];
    const stockQuantity = data[i][5];

    if (stockQuantity > 100) {
      status = 'In Stock';
    }

    else if(stockQuantity <= 100 && stockQuantity > 0){
      status = 'Repurchase Needed';
      sendStockNotification(itemID, itemName, stockQuantity);
    }

    else if(stockQuantity === 0){
      status = 'Sold Out';
      sendStockNotification(itemID, itemName, stockQuantity);
    }

    sheet.getRange(i+1, 7).setValue(status);
  }
}

function sendStockNotification(itemID, itemName, stockQuantity) {
  const recipientEmail = 'teochris423@gmail.com'; 
  const subject = `Stock Alert: ${itemName}`;
  const body = 'Dear Team,\n\n' +
               'This is to inform you that the stock levels for the following item are critically low:\n\n' +
               '- Item Name: ' + itemName + '\n' +
               '- Item ID: ' + itemID + '\n' +
               '- Current Stock Quantity: ' + stockQuantity + '\n\n' +
               'Please initiate the necessary repurchase procedures at your earliest convenience to ensure we maintain adequate stock levels.\n\n'


  MailApp.sendEmail(recipientEmail, subject, body);
}

function deductFaultyGoods(purchaseID, itemID, quantity, orderDate, supplierID, faultyAmount, receivedDate) {
  const faultysheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName('FaultyGoodsRecord');
  const itemsheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName('ItemDetails');
  const suppliersheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName('SupplierData');
  const purchasesheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName('PurchaseRecord');

  var recorddata = faultysheet.getDataRange().getValues();
  var itemdata = itemsheet.getDataRange().getValues();
  var supplierData = suppliersheet.getDataRange().getValues();
  var purchasedata = purchasesheet.getDataRange().getValues();
  
  // Update FaultyGoodsRecord sheet
  for (let i = 0; i < recorddata.length; i++) {
    if (recorddata[i][0] === purchaseID) {
      faultysheet.getRange(i + 1, 2).setValue(receivedDate); 
      faultysheet.getRange(i + 1, 3).setValue(faultyAmount); 
      break;
    }
  }

  // Update ItemDetails sheet
  for (let j = 0; j < itemdata.length; j++) {
    if (itemdata[j][0] === itemID) {
      var oldamount = itemdata[j][5];
      var newamount = oldamount + quantity - faultyAmount;
      itemsheet.getRange(j + 1, 6).setValue(newamount); 
      break;
    }
  }

  // Update SupplierData sheet
  for (let k = 0; k < supplierData.length; k++) {
    if (supplierData[k][0] === supplierID) {
      var oldfaulty = supplierData[k][11];
      var newaveragefaulty = (oldfaulty + faultyAmount) / 2;
      suppliersheet.getRange(k + 1, 12).setValue(newaveragefaulty); 
      break;
    }
  }

  //Update PurchaseRecord sheet
  for(let l = 0; l < purchasedata.length; l++)
  {
    if(purchasedata[l][0] === purchaseID)
    {
      purchasesheet.getRange(l+1, 10).setValue(true);
      break;
    }
  }

  checkStockLevels();
}

