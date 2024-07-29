function updatePaymentDate(invoiceNumber, paymentDate) {
  var invoicesheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName("InvoiceData");
  var invoicedata = invoicesheet.getDataRange().getValues();

  for (let i = 0; i < invoicedata.length; i++) {
    if (invoicedata[i][1] === invoiceNumber) {
      invoicesheet.getRange(i + 1, 8).setValue(paymentDate);
      break;
    }
  }

  Logger.log('Invoice Number: ' + invoiceNumber + ' Payment Date: ' + paymentDate);
}

function sendEmailNotification() {
  var sheet = SpreadsheetApp.openById(salesDatabaseID).getSheetByName("Form Responses"); 
  var lastRow = sheet.getLastRow();
  var row = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  var customerEmail = row[3]; 
  var customerName = row[2]; 
  var invoiceNumber = row[1]; 
  var paymentAmount = row[4]; 
  var paymentDate = new Date(row[0]); 
  var fileurl = row[5];

  var formattedPaymentDate = Utilities.formatDate(paymentDate, Session.getScriptTimeZone(), "dd/MM/yyyy");

  var subject = 'Payment Received for Invoice ' + invoiceNumber;
  var body = 'Dear ' + customerName + ',\n\n' +
             'We have received your payment for Invoice Number ' + invoiceNumber + '.\n\n' +
             'Details:\n' +
             '- Payment Amount: RM' + paymentAmount + '\n' +
             '- Payment Date: ' + formattedPaymentDate + '\n\n' +
             'Your payment is currently under review. We will issue your receipt as soon as possible. \n\n' +
             'Thank you for your payment.\n\n' +
             'Best regards,\n' +
             'Your Company Name';
  
  MailApp.sendEmail({
    to: customerEmail,
    subject: subject,
    body: body
  })

  var adminsubject = 'Payment Proof for Invoice ' + invoiceNumber;
  var adminbody = 'Dear Admin,\n\n' +
             'A new payment has been received. Here are the details:\n\n' +
             'Date Time: ' + paymentDate + '\n' +
             'Invoice Number: ' + invoiceNumber + '\n' +
             'Customer Name: ' + customerName + '\n' +
             'Email Address: ' + customerEmail + '\n' +
             'Payment Amount: ' + paymentAmount + '\n' +
             'Payment Proof: ' + fileurl + '\n\n' +
             'Best regards,\n' +
             'Your Company Name';

  MailApp.sendEmail({
    to: 'teochris423@gmail.com', 
    subject: adminsubject,
    body: adminbody,
  });

  Logger.log('Email sent to ' + customerEmail + ' and administrator.');

  updatePaymentDate(invoiceNumber, paymentDate);
}