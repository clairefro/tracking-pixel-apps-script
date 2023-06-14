/** Update these values */
const TRACKING_PIXEL_SERVER_ENDPOINT = "YOUR ENDPOINT"; // endpoint for your deployed tracking pixel server
const SPREADSHEET_ID = "19knIGGrOPS8MmiWKzGw0UkWduPxVofau6gTzJdaWva0";
const COL_SUBJECT = 3; // integer representing the absolute column number for email subject
const COL_BODY = 4; // integer representing the absolute column number for email body
const COL_ID = 5; // integer representing the absolute column number for placement of unqiue mail merge ID
const COL_STATUS = COL_ID + 1; // integer representing the absolute column number for mail merge status. Defaults to right of ID col

/** Expects email to column to be first column of selected range */
/** Does not account for bounces */
function sendEmailsWithTrackingPixel() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getActiveSheet();
  var rg = sh.getActiveRange();
  var data = rg.getValues();

  var activeRangeA1Notation = rg.getA1Notation();
  var startRow = getStartRow(activeRangeA1Notation);
  Logger.log("Change!");
  Logger.log(startRow);

  for (var i = 0; i < data.length; i++) {
    var recipient = data[i][0]; // Assuming recipient email addresses are in the first column of selected range

    // only run on rows that contain an email (ignore table headers)
    if (recipient.match(/@/)) {
      // Generate unique tracking pixel URL for each recipient
      const id = createID(recipient);
      var trackingPixel = `${TRACKING_PIXEL_SERVER_ENDPOINT}?id=${id}`;

      // Compose email message
      var subject = sh.getRange(startRow + i, COL_SUBJECT).getValue();
      var body =
        sh.getRange(startRow + i, COL_BODY).getValue() +
        `<img style="display:none;" src="${trackingPixel}">`;
      Logger.log(body);

      const idCell = sh.getRange(startRow + i, COL_ID);
      const statusCell = sh.getRange(startRow + i, COL_STATUS);
      try {
        // Send the email
        GmailApp.sendEmail(recipient, subject, "", { htmlBody: body });
        idCell.setValue(id);
        idCell.setBackground("#cfcfcf");
        statusCell.setValue("Sent!");
        statusCell.setBackground("#fffaa1");
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

function doGet(e) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getActiveSheet();

  var id = e.parameter.id;

  const foundCell = findCellById(sheet, id);

  if (foundCell) {
    const statusCell = sheet.getRange(foundCell.row, COL_STATUS);
    statusCell.setValue("Opened");
    statusCell.setBackground("green");
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

// function logToSpreadsheet(sheet, message) {
//   var timestamp = new Date().toISOString();
//   sheet.appendRow([timestamp, message]);
// }
