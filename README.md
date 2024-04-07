# automatic-invoice
 Automatic create invoice in Google docs after submitting google form. then auto export into PDF file and send it to recipient in google form
```ruby
///////////////////////////////////////////////////////////////////////////////////////////////
// BEGIN EDITS ////////////////////////////////////////////////////////////////////////////////

const TEMPLATE_FILE_ID = 'your template file id';
const DESTINATION_FOLDER_ID = 'destination folder id';
const CURRENCY_SIGN = 'THB ';

// END EDITSasdf //////////////////////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////////////////////////
// WARNING: EDITING ANYTHING BELOW THIS LINE WILL CHANGE THE BEHAVIOR OF THE SCRIPT. //////////
// DO SO AT YOUR OWN RISK.//// ////////////////////////////////////////////////////////////////
// ----------------------------------------------------------------------------------------- //

// Converts a float to a string value in the desired currency format
function toCurrency(num) {
    var fmt = Number(num).toFixed(2);
    // Split the string into integer and decimal parts
    var parts = fmt.split('.');

    // Add commas to the integer part
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');

    // Rejoin the parts and add the currency sign
    return `${CURRENCY_SIGN}${parts.join('.')}`;
}

// Format datetimes to: YYYY-MM-DD
function toDateFmt(dt_string) {
  var millis = Date.parse(dt_string);
  var date = new Date(millis);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);

  // Return the date in YYYY-mm-dd format
  return `${year}-${month}-${day}`;
}


// Parse and extract the data submitted through the form.
function parseFormData(values, header) {
    // Set temporary variables to hold prices and data.
    var subtotal = 0;
    var deposit = 0;
    var response_data = {};

    // Iterate through all of our response data and add the keys (headers)
    // and values (data) to the response dictionary object.
    for (var i = 0; i < values.length; i++) {
      // Extract the key and value
      var key = header[i];
      var value = values[i];

      // If we have a price, add it to the running subtotal and format it to the
      // desired currency.
      if (key.toLowerCase().includes("price") && value && value !== 0) {
        response_data[key] = value; // Add only existing prices
        subtotal += value;
        value = toCurrency(value);

    
      // If there is a deposit, track it so we can adjust the total later and
      // format it to the desired currency.
      } else if (key.toLowerCase().includes("deposit")) {
        deposit += value;
        value = toCurrency(value);
      
      // Format dates
      } else if (key.toLowerCase().includes("date")) {
        value = toDateFmt(value);
      }

      // Add the key/value data pair to the response dictionary.
      response_data[key] = value;
    }

    // Once all data is added, we'll adjust the subtotal and total
    response_data["sub_total"] = toCurrency(subtotal);
    response_data["total"] = toCurrency(subtotal - deposit);

    return response_data;
}

// Helper function to inject data into the template
function populateTemplate(document, response_data) {

    // Get the document header and body (which contains the text we'll be replacing).
    var document_header = document.getHeader();
    var document_body = document.getBody();

    // Replace variables in the header
    for (var key in response_data) {
      var match_text = `{{${key}}}`;
      var value = response_data[key];

      // Replace our template with the final values
      document_header.replaceText(match_text, value);
      document_body.replaceText(match_text, value);
    
    }
}
//sending email function
function sendDocumentByEmail(documentId, emailAddress,recipientPreFix,recipientFirstName,recipientLastName) {
  var subject = 'subject of email';

  // Open the Google Docs document using its ID
  var doc = DocumentApp.openById(documentId);

  // Convert the Google Docs document to PDF
  var pdfBlob = doc.getAs(MimeType.PDF);

  //getAs('application/pdf')
  //var pdfBlob = DocumentApp.openById(document.getId()).getAs(MimeType.PDF);;
  var message = 'เรียน '+ recipientPreFix+' '+recipientFirstName+' '+recipientLastName+ '\n'+'\tใบเสร็จรับชำระเงินค่ะ '+'\n\n'+'ขอบคุณค่ะ'+'\n\n'+'ติดตามข่าวสารอัพเดทกิจกรรม หรือทริปต่างๆได้ที่'+'\n'+'Facebook Page : Your facebook page'+'\n'+'Line Official : @yourlineofficial (มี @)';

  //add picture into email
  var picture1 = DriveApp.getFileById('image ID'); //public with link
  var inlineImages = {};
  inlineImages[picture1.getId()] = picture1.getBlob();

  // HTML body with logo embedded
  var htmlBody = 
    '<p>เรียน ' + recipientPreFix + ' ' + recipientFirstName + ' ' + recipientLastName + ',</p>' +
    '<p>&nbsp;&nbsp;&nbsp;&nbsp;ใบเสร็จรับชำระเงินค่ะ </p>' +
    '<p>ขอบคุณค่ะ</p>' +
    '<p>ติดตามข่าวสารอัพเดทกิจกรรม หรือทริปต่างๆได้ที่</p>' +
    '<p>Facebook Page : Your facebook page<br>' +
    'Line Official : @yourlineofficial (มี @)</p>'+
    '<p><img src="cid:' + picture1.getId() + '"></p>';

  // Send email with the document attached
  GmailApp.sendEmail(emailAddress, subject, message, {
    attachments: [pdfBlob],
    htmlBody: htmlBody,
    inlineImages : inlineImages
  });
}


// Function to populate the template form
function createDocFromForm() {

  // Get active sheet and tab of our response data spreadsheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  var last_row = sheet.getLastRow() - 1;

  // Get the data from the spreadsheet.
  var range = sheet.getDataRange();
 
  // Identify the most recent entry and save the data in a variable.
  var data = range.getValues()[last_row];
  
  // Extract the headers of the response data to automate string replacement in our template.
  var headers = range.getValues()[0];

  // Parse the form data.
  var response_data = parseFormData(data, headers);

  // Retreive the template file and destination folder.
  var template_file = DriveApp.getFileById(TEMPLATE_FILE_ID);
  // target folder is not necessary if you not define this. document will keep in destination folder above.
  var target_folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);

  // Copy the template file so we can populate it with our data.
  // The name of the file will be the invoice number and The first name of customer and date in the format: number_firstname_date
  var filename = `${response_data["autonumber"]}_${response_data["First Name"]}_${response_data["Receipt Date"]}`;
  var document_copy = template_file.makeCopy(filename, target_folder);

  // Open the copy.
  var document = DocumentApp.openById(document_copy.getId());
  var documentId = document_copy.getId()

  // Populate the template with our form responses and save the file.
  populateTemplate(document, response_data);
  document.saveAndClose();

  // Introduce a delay (e.g., 1 minute) before sending the email
  Utilities.sleep(30000); // half minute delay

  // Get the email address,prefix,first name,last name from the form response.
  var emailAddress = response_data['Email'];
  var recipientPreFix = response_data['Prefix'];
  var recipientFirstName = response_data['First Name'];
  var recipientLastName = response_data['Last Name'];

  // Send the document to the email address
  sendDocumentByEmail(documentId, emailAddress,recipientPreFix,recipientFirstName,recipientLastName);
    
}

```
