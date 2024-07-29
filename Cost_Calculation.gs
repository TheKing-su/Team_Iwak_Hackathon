function costCalculation(supplierID)
{
  var databaseID = "1rIoeWrQI6NvJategdgNaPrITlr5FV6PrgVfThsVW13s";
  var purchaseSheet = SpreadsheetApp.openById(databaseID);
  let purchaseDetails = purchaseSheet.getSheetByName('PurchaseRecord');
  let supplierDetails = purchaseSheet.getSheetByName('SupplierData');
  let itemDetails = purchaseSheet.getSheetByName('ItemDetails');

  var purchaseData = purchaseDetails.getRange('A:I').getValues();
  
  for(i = 0; i < purchaseData.length; i++)
  {
    if(purchaseData[i][5] === supplierID)
    {
      var itemID = purchaseData[i][1];
      var purchaseUnitPrice = purchaseData[i][2];
      var purchaseQuantity = purchaseData[i][3];
      var supplierID = purchaseData[i][5];
      var shippingCost = purchaseData[i][6];
      var handlingCost = purchaseData[i][7];
      var storageCost = purchaseData[i][8];

      Logger.log("Selected purchase record: " + purchaseData[i][0]);
    }
  }

  if(!supplierID)
  {
    Logger.log("Invalid Supplier ID");
  }

  var itemData = itemDetails.getRange('A:G').getValues();

  for(j = 0; j < itemData.length; j++)
  {
    if(itemData[j][0] === itemID)
    {
      var sellingUnitPrice = itemData[j][4];

      Logger.log("Selected Item ID: " + itemData[j][0])
    }
  }

  if(!itemID)
  {
    Logger.log("Ivalid Item ID");
  }

  //COGS
  var calculateCOGS = purchaseUnitPrice * purchaseQuantity

  //PPV
  var calculatePPV = (purchaseUnitPrice - sellingUnitPrice) * purchaseQuantity;

  //TCO
  var calculateTCO = calculateCOGS + shippingCost + handlingCost + storageCost;

  Logger.log("COGS: " + (calculateCOGS));
  Logger.log("PPV: " + (calculatePPV));
  Logger.log("TCO: "  + (calculateTCO));

  var supplierData = supplierDetails.getDataRange().getValues();

  for(k = 0; k < supplierData.length; k++)
  {
    if(supplierData[k][0] === supplierID)
    {
      supplierDetails.getRange('I'+ (k+1)).setValue(calculateCOGS);
      supplierDetails.getRange('J'+ (k+1)).setValue(calculatePPV);
      supplierDetails.getRange('K'+ (k+1)).setValue(calculateTCO);
    }
  }

  if(!supplierID)
  {
    Logger.Log("Invalid Supplier ID");
  }
}