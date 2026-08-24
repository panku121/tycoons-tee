/**
 * The Tycoon's - Order Form
 *
 * Saving this file is not enough. Submissions are written to the "tycoons"
 * tab in the tycoonsPPL spreadsheet. The published web app continues to run
 * the previous version until you deploy again:
 *
 * 1. Open the Google Sheet: tycoonsPPL
 * 2. Extensions > Apps Script
 * 3. Replace the script with this file and click Save
 * 4. Deploy > Manage deployments
 * 5. Edit the existing Web app
 *    (or choose Deploy > New deployment > Web app on first setup)
 * 6. Version: New version
 * 7. Deploy
 * 8. If the Web App URL is unchanged, leave js/script.js as it is
 *
 * If you created a new deployment and the URL changed, paste the new URL
 * into GOOGLE_SCRIPT_URL in js/script.js.
 */

var SHEET_NAME = "tycoons";
var HEADERS = [
  "Timestamp",
  "Full Name",
  "Phone Number",
  "T-Shirt Size",
  "Sleeve Length",
  "Jersey Number",
  "Jersey Name",
  "Trouser Size"
];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet = getOrCreateSheet();
    ensureHeaders(sheet);

    var data = getRequestData(e);
    if (!data.fullName && !data.phone) {
      return jsonResponse({ result: "ready", message: "The Tycoon's form connected." });
    }

    writeOrderRow(sheet, data);
    return jsonResponse({ result: "success" });
  } catch (error) {
    return jsonResponse({ result: "error", message: String(error) });
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function ensureHeaders(sheet) {
  migrateSleeveLengthColumn(sheet);
  migrateRemovePaymentStatusColumn(sheet);

  var firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  var headerMismatch = HEADERS.some(function (header, index) {
    return String(firstRow[index] || "").trim() !== header;
  });
  if (headerMismatch) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  var backNumberCol = columnForHeader(sheet, "Jersey Number");
  if (backNumberCol > 0) {
    sheet.getRange(2, backNumberCol, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  }
}

function migrateSleeveLengthColumn(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var hasSleeve = headers.some(function (header) {
    return String(header || "").trim() === "Sleeve Length";
  });
  if (hasSleeve) return;

  var sizeCol = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === "T-Shirt Size") {
      sizeCol = i + 1;
      break;
    }
  }
  if (sizeCol > 0) {
    sheet.insertColumnAfter(sizeCol);
    sheet.getRange(1, sizeCol + 1).setValue("Sleeve Length").setFontWeight("bold");
  }
}

function migrateRemovePaymentStatusColumn(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === "Payment Status") {
      sheet.deleteColumn(i + 1);
      return;
    }
  }
}

function writeOrderRow(sheet, data) {
  var nextRow = Math.max(sheet.getLastRow() + 1, 2);
  var backNumber = String(data.backNumber || "");
  var lowerSize = data.lowerSize || "";
  var sleeveLength = data.sleeveLength || "";
  var backNumberCol = columnForHeader(sheet, "Jersey Number");
  var rowValues = [
    new Date(),
    data.fullName || "",
    data.phone || "",
    data.size || "",
    sleeveLength,
    backNumber,
    data.backName || "",
    lowerSize
  ];

  sheet.getRange(nextRow, 1, 1, rowValues.length).setValues([rowValues]);
  if (backNumberCol > 0) {
    sheet.getRange(nextRow, backNumberCol).setNumberFormat("@").setValue(backNumber);
  }
  setCellByHeader(sheet, nextRow, "Sleeve Length", sleeveLength);
  setCellByHeader(sheet, nextRow, "Trouser Size", lowerSize);
}

function setCellByHeader(sheet, row, headerName, value) {
  var lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var col = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === headerName) {
      col = i + 1;
      break;
    }
  }
  if (col < 1) {
    col = HEADERS.indexOf(headerName) + 1;
  }
  if (col > 0) {
    sheet.getRange(row, col).setValue(value);
  }
}

function columnForHeader(sheet, headerName) {
  var lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim() === headerName) {
      return i + 1;
    }
  }
  return HEADERS.indexOf(headerName) + 1;
}

function getRequestData(e) {
  var params = {};
  copyFields(params, e && e.parameter);
  copyFields(params, firstValues(e && e.parameters));

  if (e && e.postData && e.postData.contents) {
    var type = String(e.postData.type || "");
    if (type.indexOf("json") !== -1 || type.indexOf("text/plain") !== -1) {
      try {
        copyFields(params, JSON.parse(e.postData.contents));
      } catch (err) {}
    }
  }

  return {
    fullName: params.fullName || "",
    phone: params.phone || "",
    size: params.size || "",
    sleeveLength: params.sleeveLength || "",
    backNumber: params.backNumber || "",
    backName: params.backName || "",
    lowerSize: params.lowerSize || ""
  };
}

function copyFields(target, source) {
  if (!source) return;
  Object.keys(source).forEach(function (key) {
    var value = source[key];
    if (value === undefined || value === null || value === "") return;
    target[key] = value;
  });
}

function firstValues(parameters) {
  if (!parameters) return {};
  var result = {};
  Object.keys(parameters).forEach(function (key) {
    var value = parameters[key];
    result[key] = Array.isArray(value) ? value[0] : value;
  });
  return result;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
