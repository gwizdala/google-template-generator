/**
 * Template Generator Utilities
 * Helper functions used by Main
 * 
 * @author: @gwizdala
 * 
 */

//// CONSTANTS
// "Set it and forget it" values. Update these to match your setup
// The statuses that are output to the spreadsheet to indicate progress processing the request
const GENERATION_STATUS = {
  GENERATING: "Generating",
  SUCCESS: "Success",
  ERROR: "Error"
};

//// HELPERS
// PROCESSING HELPERS
/**
 * Calculates duration between two dates in days
 * 
 * @param {string} startDate the starting date
 * @param {string} endDate the ending date
 * @return {string} the duration between the two dates in days, or "TBD"
 */
function calculateDuration(startDate, endDate) {
  if (Date.parse(startDate) && Date.parse(endDate)) {
    const date1 = new Date(startDate);
    const date2 = new Date(endDate);
    const timeDiff = date2.getTime() - date1.getTime();
    if (timeDiff > 0) {
      return `${Math.round(timeDiff / (1000 * 3600 * 24))} Days`;
    }
  }

  return "TBD";
}

// EMAIL HELPERS
/**
 * Sends a success email based on the SuccessMessage email template
 * 
 * @param {string} email the user's email
 * @param {Array[object]} files an array of the name/link of the files to generate
 * @param {string} companyName the name of the company
 */
function sendSuccessEmail(email, files, companyName) {
  const now = new Date();
  const generatedAt = `${now.toISOString()}`;
  const subject = 'Template Generation Successful';
  
  var emailTemplate = HtmlService.createTemplateFromFile('SuccessMessage');
  emailTemplate.data = {
    files: files,
    companyName,
    generatedAt,
    title: subject
  };

  Logger.log(`success email sent to ${JSON.stringify(email)}`);
  MailApp.sendEmail({
    to: email,
    noReply: true,
    subject: subject,
    htmlBody: emailTemplate.evaluate().getContent()
  });
}

/**
 * Sends a error email based on the ErrorMessage email template
 * 
 * @param {string} errorMessage the error that occured during generation
 */
function sendErrorEmail(errorMessage) {
  const now = new Date();
  const generatedAt = `${now.toISOString()}`;
  const subject = 'Template Generation Failed';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetLink = spreadsheet.getUrl();
  const editors = spreadsheet.getEditors();
  const editorEmails = editors.map(user => user.getEmail()).join(',');
  
  var emailTemplate = HtmlService.createTemplateFromFile('ErrorMessage');
  emailTemplate.data = {
    errorMessage,
    sheetLink,
    generatedAt,
    title: subject
  };

  Logger.log(`error email sent to ${JSON.stringify(editorEmails)}`);
  MailApp.sendEmail({
    to: editorEmails,
    noReply: true,
    subject: subject,
    htmlBody: emailTemplate.evaluate().getContent()
  });
}

// SHEET HELPERS
/**
 * Convert rows in a given sheet to an object.
 * 
 * @param {Spreadsheet} spreadsheet The the spreadsheet we are pulling this data from
 * @param {string} sheetName The name of the tab/sheet to retrieve data from.
 * @returns {Object<Object>|null} An object of objects, keyed by the first column
 * Returns null if the sheet is not found or no data.
 */
function getSheetDataAsObjects(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Sheet '${sheetName}' not found in the active spreadsheet.`);
    return null;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length === 0 || values[0].length === 0) {
    Logger.log(`No data or headers found in sheet '${sheetName}'.`);
    return []; // Return an empty array if no data or headers
  }

  // The first row contains the headers
  // Convert the header row to camelCase
  const headers = values[0].map(value => 
    value.toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-zA-Z0-9 ]/g, ' ')
    .split(' ')
    .map((word, index) => {
      if (index === 0) {
        return word.toLowerCase();
      }
      return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    })
    .join('')
  );

  // The rest of the rows are the data
  const dataRows = values.slice(1);

  const result = {};

  dataRows.forEach(row => {
    const rowObject = {};
    for (let i = 1; i < headers.length; i++) {
      rowObject[headers[i]] = row[i];
    }
    result[row[0]] = rowObject;
  });

  return result;
}

/**
 * Sets a value on a current row and cell
 * 
 * @param {SpreadsheetApp} sheet the sheet where the value is going to be set
 * @param {integer} rowIndex the row (0-indexed) where the value should be set
 * @param {integer} columnIndex the column (0-indexed) where the value should be set
 * @param {string} value the value that should be set in the cell
 */
function setCellValue(sheet, rowIndex, columnIndex, value) {
  sheet.getRange(rowIndex, columnIndex+1).setValue(value);
}