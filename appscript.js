/** Update these values */
const TRACKING_PIXEL_SERVER_ENDPOINT = "xxxxxxx"; // endpoint for your deployed tracking pixel server
const EMAIL_VALIDATION_SERVER_ENDPOINT = "xxxxxxxxx"; // endpoint for server that validates email format and mx. must accept query param "email"
const SPREADSHEET_ID = "xxxxxx";
const COL_SUBJECT = 3; // integer representing the absolute column number for email subject
const COL_BODY = 4; // integer representing the absolute column number for email body
const COL_ID = 5; // integer representing the absolute column number for placement of unqiue mail merge ID
const COL_STATUS = COL_ID + 1; // integer representing the absolute column number for mail merge status. Defaults to right of ID col

/** Expects email to column to be first column of selected range */
function sendEmailsWithTrackingPixel() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();
  var rg = sh.getActiveRange();
  var data = rg.getValues();

  var activeRangeA1Notation = rg.getA1Notation();
  var startRow = getStartRow(activeRangeA1Notation);

  for (var i = 0; i < data.length; i++) {
    var recipient = data[i][0]; // Assuming recipient email addresses are in the first column of selected range

    // only run on rows that contain an email (ignore table headers)

    if (recipient.match(/@/)) {
      // Generate unique tracking pixel URL for each recipient
      const id = createID(recipient);
      // get status column cell
      const statusCell = sh.getRange(startRow + i, COL_STATUS);

      /** 1. Check if email address is valid */
      // check that email is valid. Only send to valid emails
      var evURL = `${EMAIL_VALIDATION_SERVER_ENDPOINT}?email=${recipient}`;

      var response = UrlFetchApp.fetch(evURL, {
        method: "GET",
        muteHttpExceptions: true,
      });
      var evStatusCode = response.getResponseCode();
      var evContent = response.getContentText();

      if (evStatusCode === 200) {
        const isValid = JSON.parse(evContent).valid;
        if (!isValid) {
          // email bounced
          statusCell.setValue("BOUNCED");
          statusCell.setBackground("orange");
        } else {
          /** 1. Build and send email */
          // Generate tracking pixel for email body
          var trackingPixel = `${TRACKING_PIXEL_SERVER_ENDPOINT}?id=${id}`;

          // Compose email message
          var subject = sh.getRange(startRow + i, COL_SUBJECT).getValue();
          var body =
            sh.getRange(startRow + i, COL_BODY).getValue() +
            `<img style="display:none;" src="${trackingPixel}">`;
          Logger.log(body);

          const idCell = sh.getRange(startRow + i, COL_ID);
          try {
            // Send the email
            GmailApp.sendEmail(recipient, subject, "", { htmlBody: body });
            idCell.setValue(id);
            idCell.setBackground("#cfcfcf");
            statusCell.setValue("SENT");
            statusCell.setBackground("#fffaa1");
            statusCell.setNote(`Sent: ${timestampISO()}`);
          } catch (error) {
            // Handle the error
            Logger.log(
              `An error occurred when sending to '${recipient}': ${error}`
            );
            statusCell.setValue("ERR");
            statusCell.setBackground("red");
          }
        }
      }
    } else {
      statusCell.setValue(
        "ERROR when attempting to validate email. Not sent - Try again"
      );
      statusCell.setBackground("yellow");
      Logger.log("Error:", evContent, evContent);
    }
  }
}

function getStartRow(a1range) {
  var regex = /^\D+(\d+)/;
  var match = a1range.match(regex);
  var rowNumber = parseInt(match[1]);
  if (rowNumber) {
    return rowNumber;
  } else {
    return null;
  }
}

function createID(email) {
  var timestamp = new Date().getTime().toString();
  return Sha256Hash(email + timestamp);
}

function Sha256Hash(value) {
  return BytesToHex(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value)
  );
}

function BytesToHex(bytes) {
  let hex = [];
  for (let i = 0; i < bytes.length; i++) {
    let b = parseInt(bytes[i]);
    if (b < 0) {
      c = (256 + b).toString(16);
    } else {
      c = b.toString(16);
    }
    if (c.length == 1) {
      hex.push("0" + c);
    } else {
      hex.push(c);
    }
  }
  return hex.join("");
}

/** Webhook */
function doGet(e) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getActiveSheet();

  var id = e.parameter.id;

  const foundCell = findCellById(sheet, id);

  if (foundCell) {
    const statusCell = sheet.getRange(foundCell.row, COL_STATUS);
    statusCell.setValue("OPENED");
    statusCell.setBackground("green");
    const oldNote = statusCell.getNote();
    statusCell.setNote(`${oldNote}\nOpened: ${timestampISO()}`);
  }

  // Return the response
  return ContentService.createTextOutput("ok");
}

function findCellById(sheet, id) {
  var columnRange = sheet.getRange(1, COL_ID, sheet.getLastRow(), 1);

  // Create a text finder and search for the value in the column range
  var textFinder = columnRange.createTextFinder(id);
  var foundCells = textFinder.findAll();

  if (foundCells.length) {
    // Text is found in the range. ID is unique so should only be one record
    var cell = foundCells[0];
    var row = cell.getRow();
    var col = cell.getColumn();
    return { row, col };
  }
}

function timestampISO() {
  return new Date().toISOString();
}

/** For logging webhook calls, because google apps script is weird  */

// function logToSpreadsheet(sheet, message) {
//   var timestamp = new Date().toISOString();
//   sheet.appendRow([timestamp, message]);
// }
